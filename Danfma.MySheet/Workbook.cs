using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// A user-supplied function implementation. Receives the raw (unevaluated) arguments so it can decide
/// what to evaluate (lazy/short-circuit) and the workbook for context. Returns a <see cref="ComputedValue"/>
/// — a scalar (<c>double</c>/<c>bool</c>/<c>string</c>) converts implicitly; use <c>ComputedValue.Blank</c>
/// or <c>ComputedValue.Error(...)</c> for the rest.
/// </summary>
public delegate ComputedValue CustomFunction(Expression[] arguments, Workbook workbook);

[MemoryPackable]
public sealed partial class Workbook
{
    // A registered custom function paired with its (optional) declared arity. MinArgs/MaxArgs default to
    // "unconstrained" (0.. int.MaxValue) when a caller does not declare them, so an un-annotated registration
    // behaves exactly as before this type existed -- see RegisterFunction's defaults.
    internal readonly record struct CustomFunctionEntry(
        CustomFunction Function,
        int MinArgs,
        int MaxArgs
    );

    // Host-registered custom functions; not serialized (behavior is re-registered after deserialization).
    [MemoryPackIgnore]
    private Dictionary<string, CustomFunctionEntry>? _functions;

    // Memoized cell values; not serialized. Invalidation is explicit (see InvalidateCache). The dense paged
    // store addresses cells numerically (sheet handle + col/row derived on the fly from the A1 id), replacing
    // the old ConcurrentDictionary<(string,string), ComputedValue> — no per-cell box and no per-lookup string
    // tuple hash. Lazily created race-free (never `= new()` on the field: MemoryPack bypasses initializers).
    [MemoryPackIgnore]
    private SheetValueStore? _valueStore;

    // The value-store geometry this workbook was constructed with (page/group sizes, sparsity thresholds).
    // Runtime CONFIG, not document state: [MemoryPackIgnore], so the wire schema is untouched and it comes back
    // null after a Load — the ValueStore accessor falls back to ValueStoreOptions.Default in that case (the
    // field-initializer bypass lesson: never rely on `= ...` here). Captured once at construction; immutable.
    [MemoryPackIgnore]
    private ValueStoreOptions? _valueStoreOptions;

    /// <summary>Creates a workbook with the default runtime options. This is also the constructor MemoryPack
    /// uses to materialize a deserialized workbook (which then falls back to the default value-store options,
    /// since the options field is runtime config and is never serialized).</summary>
    [MemoryPackConstructor]
    public Workbook() { }

    /// <summary>
    /// Creates a workbook with explicit runtime <paramref name="options"/> (currently the dense value store's
    /// page/group sizes and sparsity thresholds). The options are validated here — a non-power-of-two size or an
    /// out-of-range value throws <see cref="ArgumentException"/> — and captured immutably; they apply to value
    /// stores created after construction (i.e. from the first evaluation). Options are runtime configuration and
    /// are never serialized, so the file format is unchanged.
    /// </summary>
    public Workbook(WorkbookOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        options.ValueStore.Validate();
        _valueStoreOptions = options.ValueStore;
    }

    private SheetValueStore ValueStore
    {
        get
        {
            var existing = _valueStore;
            if (existing is not null)
            {
                return existing;
            }

            var created = new SheetValueStore(_valueStoreOptions ?? ValueStoreOptions.Default);
            return Interlocked.CompareExchange(ref _valueStore, created, null) ?? created;
        }
    }

    // Test hook: the store instance (creating it if needed), so a test can assert a configured geometry flowed
    // through. Not part of the public API.
    internal SheetValueStore ValueStoreForTesting => ValueStore;

    // The dense value store, exposed to the range enumerator (RangeValueSequence) so a hot fold takes the HIT
    // path — a lock-free dense read — inline, without re-entering GetCellValueDense (and its ValueStore getter)
    // per cell. A MISS still routes back through GetCellValueDense for the cycle guard + on-demand evaluation,
    // so the semantics are identical; only the already-memoized read is shortened. Internal fast path.
    internal SheetValueStore DenseStore => ValueStore;

    // Cells currently being evaluated on the calling thread, to detect circular references. Thread-local so
    // concurrent (and benign) re-evaluation of the same cell on different threads is not a false cycle — the
    // reason this stays a per-thread set and NOT a shared per-slot bit (a shared bit would let one thread's
    // in-flight mark spuriously trip another's read as a cycle). The key is the numeric address (handle, col,
    // row), which drops the old string-tuple hashing; the rare non-A1 id falls back to _evaluatingOverflow.
    [ThreadStatic]
    private static HashSet<(int Handle, int Col, int Row)>? _evaluating;

    // Cycle guard for the non-A1 overflow path (a host reading a cell under a non-address key). Dormant in
    // normal use; keeps the exact enter/exit semantics for that path too.
    [ThreadStatic]
    private static HashSet<(string Sheet, string Id)>? _evaluatingOverflow;

    // Sheet names are case-insensitive, like Excel.
    public ConcurrentDictionary<string, Sheet> Sheets { get; private set; } =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Workbook-level defined names (case-insensitive, like Excel), each mapping to the
    /// <see cref="Expression"/> it stands for — typically a sheet-qualified range or cell, but any
    /// expression (a constant, a formula, another name) is allowed. Define them through
    /// <see cref="DefineName(string, Expression)"/> / <see cref="DefineName(string, string)"/>; a
    /// <see cref="Expressions.NameReference"/> in a formula resolves against this map (after the LET scope).
    /// Mutating this dictionary directly (rather than through <see cref="DefineName(string, Expression)"/>)
    /// is NOT tracked by <see cref="NamesVersion"/>, so a <see cref="RecalculationEngine"/> built over the
    /// workbook would not notice the change — go through <see cref="DefineName(string, Expression)"/> instead.
    /// </summary>
    // MemoryPack serializes members in declaration order; this MUST stay the LAST serialized member of
    // Workbook so the schema is append-only — files written before it existed (which carry only Sheets)
    // still load, leaving this empty.
    public Dictionary<string, Expression> DefinedNames { get; private set; } =
        new(StringComparer.OrdinalIgnoreCase);

    // Monotonic counter of DefinedNames mutations, bumped by DefineName. A defined name is resolved into the
    // reverse dependency graph at BUILD time (DependencyExtractor.ResolveName bakes in the definition current at
    // that point), so redefining a name — repointing it at a different reference — changes the dependency
    // structure exactly like a formula edit does, but touches no Sheet and so would never bump any
    // Sheet.StructuralVersion. RecalculationEngine snapshots this alongside the per-sheet versions to detect
    // that staleness. Runtime-only ([MemoryPackIgnore]): a loaded workbook starts at 0 and any engine built
    // after Load rebuilds from the current state anyway. Single-thread edit contract, same as StructuralVersion.
    [MemoryPackIgnore]
    private long _namesVersion;

    /// <summary>The count of <see cref="DefineName(string, Expression)"/> calls this workbook has seen; the
    /// reverse dependency graph uses it (together with each sheet's <see cref="Sheet.StructuralVersion"/>) to
    /// detect that it went stale. Internal — part of the recalculation contract, not host API.</summary>
    internal long NamesVersion => _namesVersion;

    // MemoryPack rebuilds the dictionaries with the default (case-sensitive) comparer, and older files
    // carry no DefinedNames at all (null after deserialization); restore ours in both cases.
    [MemoryPackOnDeserialized]
    private void RestoreComparers()
    {
        Sheets = new ConcurrentDictionary<string, Sheet>(Sheets, StringComparer.OrdinalIgnoreCase);
        DefinedNames = DefinedNames is null
            ? new Dictionary<string, Expression>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, Expression>(DefinedNames, StringComparer.OrdinalIgnoreCase);
    }

    public Sheet this[string key] => Sheets[key];

    /// <summary>
    /// Tries to resolve a sheet by name (case-insensitive, like Excel) without throwing, so a host can probe
    /// for a sheet the way a formula does. Unlike the <see cref="this[string]"/> indexer — which keeps
    /// throwing <see cref="KeyNotFoundException"/> for a direct host lookup, exactly like a dictionary — a
    /// MISSING sheet inside a formula resolves to <c>#REF!</c> (see <see cref="GetCellValue"/>) rather than
    /// aborting the evaluation.
    /// </summary>
    public bool TryGetSheet(string name, [MaybeNullWhen(false)] out Sheet sheet) =>
        Sheets.TryGetValue(name, out sheet);

    /// <summary>
    /// Returns the memoized <see cref="ComputedValue"/> of a cell — the cache stores the struct inline, so a
    /// cell referenced by many formulas is computed once with no per-cell heap box. The cache is NOT
    /// invalidated automatically on mutation — call <see cref="InvalidateCache"/> after editing cells.
    /// </summary>
    public ComputedValue GetCellValue(string sheetName, string id)
    {
        var store = ValueStore;

        // Derive the numeric address once (no-alloc). Only a direct host call can produce a non-A1 id — every
        // formula reference is normalized to A1 by the parser — so that rare case takes the overflow path,
        // which preserves the old string-keyed dictionary behavior exactly.
        if (!CellAddress.TryGetColumnRow(id, out var column, out var row))
        {
            return GetCellValueOverflow(store, sheetName, id);
        }

        var handle = store.HandleFor(sheetName);

        if (store.TryGetDense(handle, column, row, out var cached))
        {
            return cached;
        }

        var evaluating = _evaluating ??= new();
        var key = (handle, column, row);

        if (!evaluating.Add(key))
        {
            // The cell is already on this thread's evaluation stack: a circular reference.
            return ComputedValue.Error(Error.Ref);
        }

        try
        {
            var value = EvaluateCell(sheetName, id, out var touched);
            store.SetDense(handle, column, row, value, touched);
            return value;
        }
        finally
        {
            evaluating.Remove(key);
        }
    }

    /// <summary>
    /// Resolves the dense sheet <em>handle</em> once so a range expansion can address many cells numerically
    /// without re-resolving the sheet per cell. Pass the result to <see cref="GetCellValueDense"/>. The handle
    /// is a stable index for the value store's life (it is assigned on first sight, exactly as
    /// <see cref="GetCellValue"/> does per call). Internal — part of the range-read fast path, not host API.
    /// </summary>
    internal int ResolveDenseHandle(string sheetName) => ValueStore.HandleFor(sheetName);

    /// <summary>
    /// Creates a numeric-address value reader for one sheet, resolving the sheet handle once. Bulk
    /// post-compute extraction via <c>GetCellValue(name, "C" + r)</c> pays an id-string allocation, an
    /// A1 parse and a sheet-name hash lookup PER CELL; <see cref="SheetValueReader.GetValue"/> pays
    /// none of them on a cache hit. A miss evaluates on demand with semantics IDENTICAL to
    /// <see cref="GetCellValue(string,string)"/> (memoization, cycle guard, taint propagation), so
    /// literals and never-computed formulas are served too — this is a faster address form of the
    /// same read, not a snapshot of "only what was already computed". Like <c>GetCellValue</c>, the
    /// cache is not invalidated automatically: call <see cref="InvalidateCache"/> after edits.
    /// </summary>
    public SheetValueReader GetValueReader(string sheetName)
    {
        ArgumentNullException.ThrowIfNull(sheetName);

        return new SheetValueReader(this, ResolveDenseHandle(sheetName), sheetName);
    }

    /// <summary>
    /// Spike (dirty-graph): evicts one dense cell from the memoized cache so the next read recomputes it —
    /// the selective invalidation the evict-and-pull path uses instead of <see cref="InvalidateCache"/>.
    /// No-op if the value store has not been created yet. Internal — not host API.
    /// </summary>
    internal void EvictDense(string sheetName, int column, int row)
    {
        if (_valueStore is { } store)
        {
            store.EvictDense(store.HandleFor(sheetName), column, row);
        }
    }

    /// <summary>
    /// The numeric-address twin of <see cref="GetCellValue(string,string)"/> for range expansion: given a
    /// pre-resolved sheet <paramref name="handle"/> (from <see cref="ResolveDenseHandle"/>) and a 1-based
    /// <paramref name="column"/>/<paramref name="row"/>, returns the memoized value with IDENTICAL semantics to
    /// the string path — on-demand evaluation of a miss, the per-thread cycle guard, volatile-taint propagation
    /// and the blank-at-boundary coercion. The point is the HIT: a cache hit touches no string at all (no
    /// <see cref="CellAddress.ToId"/> build, no re-parse, no per-cell sheet-handle lookup — the whole cost the
    /// range enumerators used to pay per cell). Only a MISS materializes the A1 id ONCE, purely so the cell's
    /// expression can be found and evaluated exactly as the string path does. <paramref name="sheetName"/> is
    /// still needed for that miss path (the id lookup and the evaluation context); it MUST be the same name the
    /// handle was resolved from.
    /// </summary>
    internal ComputedValue GetCellValueDense(int handle, string sheetName, int column, int row)
    {
        var store = ValueStore;

        if (store.TryGetDense(handle, column, row, out var cached))
        {
            return cached;
        }

        var evaluating = _evaluating ??= new();
        var key = (handle, column, row);

        if (!evaluating.Add(key))
        {
            // The cell is already on this thread's evaluation stack: a circular reference.
            return ComputedValue.Error(Error.Ref);
        }

        try
        {
            // Miss only: materialize the A1 id once so the expression can be located and evaluated — the same
            // work GetCellValue does, minus the per-cell string round trip a hit would have paid.
            var id = new CellAddress(column, row).ToId();
            var value = EvaluateCell(sheetName, id, out var touched);
            store.SetDense(handle, column, row, value, touched);
            return value;
        }
        finally
        {
            evaluating.Remove(key);
        }
    }

    // The non-A1 overflow path: a host may store/read a cell under a key that is not an A1 address, which the
    // dense store cannot address. It behaves exactly like the old dictionary cache (memoize, cycle-guard, taint).
    private ComputedValue GetCellValueOverflow(SheetValueStore store, string sheetName, string id)
    {
        if (store.TryGetOverflow(sheetName, id, out var cached))
        {
            return cached;
        }

        var evaluating = _evaluatingOverflow ??= new();
        var key = (sheetName, id);

        if (!evaluating.Add(key))
        {
            return ComputedValue.Error(Error.Ref);
        }

        try
        {
            var value = EvaluateCell(sheetName, id, out var touched);
            store.SetOverflow(sheetName, id, value, touched);
            return value;
        }
        finally
        {
            evaluating.Remove(key);
        }
    }

    // Evaluates one cell's expression with the cell-boundary coercion and the missing-sheet → #REF! rule,
    // reporting whether it (directly or transitively) touched a volatile. It owns ONLY the thread-local
    // volatile-taint save/reset/propagate (mirror of the _evaluating pattern): zero the flag for THIS cell,
    // evaluate, read whether a volatile was touched, then restore the caller's flag OR'd with ours so
    // volatility propagates up the evaluation stack (contagion) without a reverse dependency graph. It does NOT
    // memoize or guard cycles — the caller owns the store write and the cycle guard, addressed either densely
    // or through the overflow path.
    private ComputedValue EvaluateCell(string sheetName, string id, out bool touched)
    {
        var outerTouched = _volatileTouched;
        _volatileTouched = false;

        ComputedValue value;

        // A reference to a sheet that does not exist is a STRUCTURAL failure (#REF!), fiel ao Excel — not a
        // thrown KeyNotFoundException that would abort a whole batch. Detected here (before indexing) it covers
        // the per-cell paths: CellReference.Evaluate and the cell-by-cell range enumerators, which all funnel
        // through GetCellValue.
        if (!Sheets.TryGetValue(sheetName, out var sheet))
        {
            value = ComputedValue.Error(Error.Ref);
        }
        else
        {
            // Compute outside the store (the formula recurses back in), then the caller stores it.
            var expression = sheet[id];
            value = expression.Evaluate(new EvaluationContext(this, sheetName, id));

            // Excel parity — a formula result is NEVER blank at the CELL boundary: when a cell that HAS content
            // (its expression is not the empty BlankValue) evaluates to blank (e.g. =Sheet2!F10 with F10 empty,
            // or =IF(TRUE, F10)), Excel displays 0, so the coerced 0 is what enters the cache. A truly empty
            // cell (BlankValue expression) stays blank. The coercion is the CELL's, not the expression's:
            // Expression.Evaluate keeps blank INTERNALLY (blank still compares as ""/0/FALSE inside an
            // expression), so it is intentionally NOT touched here.
            if (expression is not BlankValue && value.Kind == ComputedValueKind.Blank)
            {
                value = ComputedValue.Number(0);
            }
        }

        touched = _volatileTouched;
        _volatileTouched = outerTouched || touched;

        return value;
    }

    /// <summary>
    /// Clears the whole memoized cache (call after editing cells) and resets the volatile epoch, so the next
    /// read re-samples the clock. Use this when inputs changed; use <see cref="Recalculate"/> for a cheap
    /// volatile-only refresh that keeps the stable cells cached.
    /// </summary>
    public void InvalidateCache()
    {
        _valueStore?.Clear();

        // The structural index is NO LONGER dropped here (3.0): it is write-maintained by the SetCell/Remove
        // choke point and lifetime-scoped, so it already reflects every cell edit. InvalidateCache clears the
        // VALUE caches (a cell's content/formula may have changed) but structure is untouched by that.

        // The value snapshots are stale once any cell changed.
        _rangeCache?.Clear();
        Volatile.Write(ref _rangeCacheEntryCount, 0);

        lock (ClockLock)
        {
            _epochNow = null;
        }
    }

    /// <summary>
    /// Advances the volatile epoch: drops from the cache only the cells that touched a volatile
    /// (<c>NOW</c>/<c>TODAY</c>/<c>RAND</c>/<c>RANDBETWEEN</c>, directly or transitively) and re-samples the
    /// clock, so the next read of those cells produces fresh values while every stable (non-volatile) cell
    /// stays cached. Does NOT recompute eagerly — values are refreshed lazily on the next read. Unlike
    /// <see cref="InvalidateCache"/>, this is the right call when only "the current time / a new random draw"
    /// should change, not the cells' inputs.
    /// </summary>
    public void Recalculate()
    {
        // Drop from the store exactly the cells that touched a volatile this epoch; they recompute lazily on
        // the next read while every stable cell stays memoized.
        _valueStore?.DropTainted();

        // The value snapshots may hold volatile-tainted values, so they are dropped on a volatile refresh too
        // (the structural index — pure structure — survives).
        _rangeCache?.Clear();
        Volatile.Write(ref _rangeCacheEntryCount, 0);

        lock (ClockLock)
        {
            _epochNow = null;
        }
    }

    /// <summary>
    /// Creates a <see cref="RecalculationEngine"/> over this workbook: a reverse dependency graph that turns a
    /// set of edited cells into just the affected cone (evict-and-pull) instead of a whole-workbook
    /// <see cref="InvalidateCache"/>, and answers which outputs a set of inputs affects. Build it AFTER the
    /// workbook is populated (typically after a first <see cref="ComputeAll"/>); it tracks structural edits so a
    /// FORMULA change rebuilds the graph transparently while value edits stay on the cheap path.
    /// </summary>
    public RecalculationEngine CreateRecalculationEngine() => new(this);

    /// <summary>
    /// Eagerly evaluates every cell in the workbook, filling the memoization cache so later reads (and a
    /// warm <see cref="Save(string, WorkbookSaveOptions)"/>) hit already-computed values instead of
    /// evaluating lazily. This is the eager counterpart to on-demand <see cref="GetCellValue"/> — the
    /// analogue of a spreadsheet engine's "calculate now". The sweep runs on a large-stack thread (via
    /// <see cref="RunWithLargeStack{T}"/>) so deep dependency chains do not overflow; the thread cost is
    /// paid once for the whole workbook, not per cell.
    /// <para>
    /// Values are memoized, so calling this on an already-warm workbook is cheap (every read is a hit).
    /// After changing inputs, call <see cref="InvalidateCache"/> first, then this, to recompute from
    /// scratch. Evaluation is on-demand recursive, so cell order does not matter — each cell is computed
    /// exactly once regardless of the sweep order.
    /// </para>
    /// </summary>
    public void ComputeAll(int stackSizeBytes = 256 * 1024 * 1024) =>
        RunWithLargeStack(
            () =>
            {
                foreach (var sheet in Sheets.Values)
                {
                    var sheetName = sheet.Name;
                    foreach (var id in sheet.Keys)
                    {
                        GetCellValue(sheetName, id);
                    }
                }

                return 0;
            },
            stackSizeBytes
        );

    /// <summary>
    /// Runs an evaluation on a thread with a large stack so very deep dependency chains (e.g. a long
    /// cumulative column) do not overflow. Wrap a whole extraction batch in one call — the thread cost
    /// is paid once, not per cell. The large stack size is a reservation; physical memory grows only
    /// with the depth actually reached.
    /// </summary>
    public static T RunWithLargeStack<T>(Func<T> work, int stackSizeBytes = 256 * 1024 * 1024)
    {
        var result = default(T)!;
        ExceptionDispatchInfo? error = null;

        var thread = new Thread(
            () =>
            {
                try
                {
                    result = work();
                }
                catch (Exception exception)
                {
                    error = ExceptionDispatchInfo.Capture(exception);
                }
            },
            stackSizeBytes
        );

        thread.Start();
        thread.Join();

        error?.Throw();

        return result;
    }

    /// <summary>
    /// Registers (or replaces) a custom function callable from formulas as <paramref name="name"/>.
    /// </summary>
    /// <param name="minArgs">
    /// Minimum argument count, inclusive. A call with fewer arguments evaluates to <c>#VALUE!</c> instead of
    /// invoking <paramref name="function"/> -- <paramref name="function"/> itself never sees an out-of-range
    /// call, so it does not need to bounds-check <c>arguments</c> before indexing it. Defaults to 0 (no
    /// minimum), which combined with the default <paramref name="maxArgs"/> reproduces the unchecked behavior
    /// from before this parameter existed.
    /// </param>
    /// <param name="maxArgs">Maximum argument count, inclusive. Defaults to <see cref="int.MaxValue"/> (no maximum).</param>
    public void RegisterFunction(
        string name,
        CustomFunction function,
        int minArgs = 0,
        int maxArgs = int.MaxValue
    ) =>
        (_functions ??= new(StringComparer.OrdinalIgnoreCase))[name] = new CustomFunctionEntry(
            function,
            minArgs,
            maxArgs
        );

    public bool TryGetFunction(string name, [MaybeNullWhen(false)] out CustomFunction function)
    {
        if (_functions is not null && _functions.TryGetValue(name, out var entry))
        {
            function = entry.Function;
            return true;
        }

        function = null;
        return false;
    }

    // FunctionCall.Evaluate's lookup: the entry carries the declared arity alongside the delegate, so the
    // out-of-range check can run BEFORE invoking the delegate (see CustomFunctionEntry). Internal: this is
    // wiring for evaluation, not something a host registering/looking up functions needs -- TryGetFunction
    // above remains the public surface for that.
    internal bool TryGetFunctionEntry(string name, out CustomFunctionEntry entry)
    {
        if (_functions is null)
        {
            entry = default;
            return false;
        }

        return _functions.TryGetValue(name, out entry);
    }

    /// <summary>
    /// Defines (or redefines) a workbook-level name that stands for <paramref name="reference"/> — usually a
    /// sheet-qualified range or cell, but any <see cref="Expression"/> is allowed (a constant, a formula,
    /// another name). Names are case-insensitive. Throws <see cref="ArgumentException"/> when the name is
    /// empty or collides with a cell-reference/boolean literal.
    /// </summary>
    public void DefineName(string name, Expression reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        NamedReferences.ValidateName(name);

        DefinedNames[name] = reference;
        unchecked
        {
            _namesVersion++;
        }
    }

    /// <summary>
    /// Convenience overload that parses <paramref name="formulaText"/> (a leading <c>=</c> is optional) into
    /// the name's expression. Because names are workbook-level with no implicit sheet, every reference in
    /// the text MUST be sheet-qualified (e.g. <c>"Data!$A$1:$A$10"</c>); an unqualified reference throws
    /// <see cref="ArgumentException"/>.
    /// </summary>
    public void DefineName(string name, string formulaText)
    {
        ArgumentNullException.ThrowIfNull(formulaText);
        NamedReferences.ValidateName(name);

        var text = formulaText.StartsWith('=') ? formulaText : "=" + formulaText;

        // Parse in a sentinel ("") sheet context: an unqualified reference then carries the empty sheet
        // name, which no real sheet has, so it is unambiguously detectable (and rejected) below.
        var expression = ExpressionParser.Parse(text, new Sheet { Name = string.Empty });

        if (HasUnqualifiedReference(expression))
        {
            throw new ArgumentException(
                $"The definition of '{name}' has an unqualified reference; defined names are workbook-level, "
                    + "so every reference must be sheet-qualified (e.g. \"Data!$A$1:$A$10\").",
                nameof(formulaText)
            );
        }

        DefinedNames[name] = expression;
        unchecked
        {
            _namesVersion++;
        }
    }

    // Walks the parsed definition for any cell/range reference left on the sentinel ("") sheet — i.e.
    // written without a sheet qualifier. FormulaWriter.Call yields a function node's argument list.
    private static bool HasUnqualifiedReference(Expression expression) =>
        expression switch
        {
            CellReference cell => cell.SheetName.Length == 0,
            RangeReference range => range.SheetName.Length == 0,
            UnionReference union => union.Areas.Any(HasUnqualifiedReference),
            BinaryOperation binary => HasUnqualifiedReference(binary.Left)
                || HasUnqualifiedReference(binary.Right),
            UnaryOperation unary => HasUnqualifiedReference(unary.Operand),
            Function function => FormulaWriter
                .Call(function)
                .Arguments.Any(HasUnqualifiedReference),
            _ => false,
        };
}
