using System.Collections;
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
    // Host-registered custom functions; not serialized (behavior is re-registered after deserialization).
    [MemoryPackIgnore]
    private Dictionary<string, CustomFunction>? _functions;

    // Memoized cell values; not serialized. Invalidation is explicit (see InvalidateCache). Stores the
    // ComputedValue struct inline — no long-lived per-cell box (the source of the Gen1 pressure the
    // ComputedValue migration removes).
    [MemoryPackIgnore]
    private ConcurrentDictionary<(string Sheet, string Id), ComputedValue>? _cache;

    // Cells currently being evaluated on the calling thread, to detect circular references. Thread-local
    // so concurrent (and benign) re-evaluation of the same cell on different threads is not a false cycle.
    [ThreadStatic]
    private static HashSet<(string Sheet, string Id)>? _evaluating;

    // === Volatile-function epoch model (F1) ===============================================================
    // A volatile cell (NOW/TODAY/RAND/RANDBETWEEN, directly or transitively) is cached WITHIN an epoch and
    // its key is recorded in _volatileTainted. Recalculate() drops just those cells and re-samples the clock
    // (epoch++); InvalidateCache() drops everything. This keeps NOW()/RAND() coherent within a pass (sampled
    // once) while staying cheap to refresh, all without a dependency graph.

    // "The cell currently being evaluated on this thread touched a volatile." Same thread-local save/reset/
    // propagate pattern as _evaluating: GetCellValue zeroes it before a cell, reads it after (to mark the
    // cell), and ORs it back into the caller's value so volatility propagates up the evaluation stack.
    [ThreadStatic]
    private static bool _volatileTouched;

    // Keys of cells that touched a volatile this epoch — a set (the byte value is unused). Concurrent because
    // evaluation can run on many threads. Lazily created race-free via the VolatileTainted accessor.
    [MemoryPackIgnore]
    private ConcurrentDictionary<(string Sheet, string Id), byte>? _volatileTainted;

    // The clock sampled once per epoch (lazily, on the first volatile read), as an Excel serial (local time).
    // null means "not yet sampled this epoch". Guarded by VolatileLock so it is sampled exactly once.
    [MemoryPackIgnore]
    private double? _epochNow;

    // Persistent RNG for RAND/RANDBETWEEN, advanced across epochs (never re-seeded per epoch, so epochs differ
    // naturally; a fixed RandomSeed makes the whole run reproducible). Created lazily under VolatileLock.
    [MemoryPackIgnore]
    private Random? _random;

    // Guards the once-per-epoch clock sampling and the not-thread-safe RNG. Lazily created (Interlocked)
    // so it survives MemoryPack deserialization, which bypasses field initializers.
    [MemoryPackIgnore]
    private object? _volatileLock;

    private object VolatileLock
    {
        get
        {
            var existing = _volatileLock;
            if (existing is not null)
            {
                return existing;
            }

            var created = new object();
            return Interlocked.CompareExchange(ref _volatileLock, created, null) ?? created;
        }
    }

    private ConcurrentDictionary<(string Sheet, string Id), byte> VolatileTainted
    {
        get
        {
            var existing = _volatileTainted;
            if (existing is not null)
            {
                return existing;
            }

            var created = new ConcurrentDictionary<(string Sheet, string Id), byte>();
            return Interlocked.CompareExchange(ref _volatileTainted, created, null) ?? created;
        }
    }

    // Backing field for the injectable clock; ignored by MemoryPack (runtime config, not persisted state).
    [MemoryPackIgnore]
    private TimeProvider? _timeProvider;

    /// <summary>
    /// The clock <c>NOW()</c>/<c>TODAY()</c> read (defaults to <see cref="TimeProvider.System"/>). Injectable
    /// so hosts can freeze time for a batch and tests can pin both the instant and the local zone. Excel uses
    /// LOCAL time, so the functions read <see cref="TimeProvider.GetLocalNow"/>. Not serialized.
    /// </summary>
    [MemoryPackIgnore]
    public TimeProvider TimeProvider
    {
        get => _timeProvider ?? TimeProvider.System;
        set => _timeProvider = value;
    }

    /// <summary>
    /// Seed for the <c>RAND</c>/<c>RANDBETWEEN</c> RNG. <c>null</c> (default) seeds it from the clock; a fixed
    /// value makes the whole run's random sequence reproducible. Set it before the first volatile read (the
    /// RNG is created lazily and never re-seeded afterwards). Not serialized (runtime config).
    /// </summary>
    [MemoryPackIgnore]
    public int? RandomSeed { get; set; }

    /// <summary>
    /// Marks the current thread's cell evaluation as having touched a volatile source. Volatile nodes call
    /// this from <c>Evaluate</c>; <see cref="GetCellValue"/> reads the flag to cache-and-mark the cell and to
    /// propagate volatility to dependents. Internal — part of the evaluation contract, not host API.
    /// </summary>
    internal void MarkVolatileTouched() => _volatileTouched = true;

    /// <summary>
    /// The current epoch's clock, sampled once (lazily) as an Excel serial from local time, so every
    /// <c>NOW()</c>/<c>TODAY()</c> in a pass agrees. Thread-safe: the first caller of the epoch samples and
    /// publishes under <see cref="VolatileLock"/>, the rest read the published value.
    /// </summary>
    internal double EpochNow()
    {
        lock (VolatileLock)
        {
            return _epochNow ??= DateSerial.FromDateTime(TimeProvider.GetLocalNow().DateTime);
        }
    }

    /// <summary>
    /// Draws the next value in <c>[0, 1)</c> from the persistent RNG (created lazily from
    /// <see cref="RandomSeed"/>) and marks the evaluation volatile. Thread-safe (<see cref="Random"/> is not).
    /// The RNG is NOT re-seeded per epoch — the sequence continues so successive epochs differ, while the
    /// per-epoch cache keeps a single cell stable within a pass.
    /// </summary>
    internal double NextRandom()
    {
        MarkVolatileTouched();

        lock (VolatileLock)
        {
            _random ??= RandomSeed is { } seed ? new Random(seed) : new Random();
            return _random.NextDouble();
        }
    }

    // Sheet names are case-insensitive, like Excel.
    public ConcurrentDictionary<string, Sheet> Sheets { get; private set; } =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Workbook-level defined names (case-insensitive, like Excel), each mapping to the
    /// <see cref="Expression"/> it stands for — typically a sheet-qualified range or cell, but any
    /// expression (a constant, a formula, another name) is allowed. Define them through
    /// <see cref="DefineName(string, Expression)"/> / <see cref="DefineName(string, string)"/>; a
    /// <see cref="Expressions.NameReference"/> in a formula resolves against this map (after the LET scope).
    /// </summary>
    // MemoryPack serializes members in declaration order; this MUST stay the LAST serialized member of
    // Workbook so the schema is append-only — files written before it existed (which carry only Sheets)
    // still load, leaving this empty.
    public Dictionary<string, Expression> DefinedNames { get; private set; } =
        new(StringComparer.OrdinalIgnoreCase);

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
        var cache = _cache ??= new();
        var key = (sheetName, id);

        if (cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var evaluating = _evaluating ??= new();

        if (!evaluating.Add(key))
        {
            // The cell is already on this thread's evaluation stack: a circular reference.
            return ComputedValue.Error(Error.Ref);
        }

        // Volatility taint (mirror of the _evaluating pattern): zero the thread-local flag for THIS cell,
        // evaluate, then see whether the cell — directly or transitively — touched a volatile. If so, cache
        // it (per-epoch cache) but record its key so Recalculate() can drop exactly those cells. Restore the
        // caller's flag OR'd with ours so volatility propagates up the evaluation stack (contagion) without a
        // reverse dependency graph.
        var outerTouched = _volatileTouched;
        _volatileTouched = false;

        try
        {
            // A reference to a sheet that does not exist is a STRUCTURAL failure (#REF!), fiel ao Excel — not
            // a thrown KeyNotFoundException that would abort a whole batch. Detected here (before indexing) it
            // covers the per-cell paths: CellReference.Evaluate and the cell-by-cell range enumerators, which
            // all funnel through GetCellValue.
            if (!Sheets.TryGetValue(sheetName, out var sheet))
            {
                var missing = ComputedValue.Error(Error.Ref);
                cache[key] = missing;
                _volatileTouched = outerTouched;

                return missing;
            }

            // Compute outside the dictionary (the formula recurses back in), then store.
            var value = sheet[id].Evaluate(new EvaluationContext(this, sheetName, id));

            var cellTouched = _volatileTouched;
            if (cellTouched)
            {
                VolatileTainted[key] = 0;
            }

            cache[key] = value;
            _volatileTouched = outerTouched || cellTouched;

            return value;
        }
        finally
        {
            evaluating.Remove(key);
        }
    }

    /// <summary>
    /// Clears the whole memoized cache (call after editing cells) and resets the volatile epoch, so the next
    /// read re-samples the clock. Use this when inputs changed; use <see cref="Recalculate"/> for a cheap
    /// volatile-only refresh that keeps the stable cells cached.
    /// </summary>
    public void InvalidateCache()
    {
        _cache?.Clear();
        _volatileTainted?.Clear();

        lock (VolatileLock)
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
        if (_volatileTainted is { } tainted)
        {
            if (_cache is { } cache)
            {
                foreach (var key in tainted.Keys)
                {
                    cache.TryRemove(key, out _);
                }
            }

            tainted.Clear();
        }

        lock (VolatileLock)
        {
            _epochNow = null;
        }
    }

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
    /// Serializes the workbook to a file (MemoryPack). The cell cache and registered custom functions are
    /// not persisted — they are rebuilt/re-registered after loading.
    /// </summary>
    public void Save(string path) => File.WriteAllBytes(path, MemoryPackSerializer.Serialize(this));

    /// <inheritdoc cref="Save"/>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = File.Create(path);
        await MemoryPackSerializer.SerializeAsync(
            stream,
            this,
            cancellationToken: cancellationToken
        );
    }

    /// <summary>Loads a workbook from a file written by <see cref="Save"/> and returns the new instance.</summary>
    public static Workbook Load(string path) =>
        MemoryPackSerializer.Deserialize<Workbook>(File.ReadAllBytes(path))
        ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");

    /// <inheritdoc cref="Load"/>
    public static async Task<Workbook> LoadAsync(
        string path,
        CancellationToken cancellationToken = default
    )
    {
        await using var stream = File.OpenRead(path);

        return await MemoryPackSerializer.DeserializeAsync<Workbook>(
                stream,
                cancellationToken: cancellationToken
            ) ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");
    }

    public void RegisterFunction(string name, CustomFunction function) =>
        (_functions ??= new(StringComparer.OrdinalIgnoreCase))[name] = function;

    public bool TryGetFunction(string name, [MaybeNullWhen(false)] out CustomFunction function)
    {
        if (_functions is null)
        {
            function = null;
            return false;
        }

        return _functions.TryGetValue(name, out function);
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
    }

    // Walks the parsed definition for any cell/range reference left on the sentinel ("") sheet — i.e.
    // written without a sheet qualifier. FormulaWriter.Call yields a function node's argument list.
    private static bool HasUnqualifiedReference(Expression expression) =>
        expression switch
        {
            CellReference cell => cell.SheetName.Length == 0,
            RangeReference range => range.SheetName.Length == 0,
            UnionReference union => union.Areas.Any(HasUnqualifiedReference),
            BinaryOperation binary =>
                HasUnqualifiedReference(binary.Left) || HasUnqualifiedReference(binary.Right),
            UnaryOperation unary => HasUnqualifiedReference(unary.Operand),
            Function function => FormulaWriter.Call(function).Arguments.Any(HasUnqualifiedReference),
            _ => false,
        };
}

public static class CollectionExtensions
{
    extension(ConcurrentDictionary<string, Sheet> sheets)
    {
        public Sheet Add(string name) =>
            sheets.GetOrAdd(name, name => new Sheet { Name = name, Index = sheets.Count });
    }
}

[MemoryPackable]
public sealed partial class Sheet : IEnumerable<KeyValuePair<string, Expression>>
{
    public required string Name { get; init; }

    // 0-based insertion order, used by the SHEET function.
    public int Index { get; init; }

    public Dictionary<string, Expression> Cells { get; init; } = new();

    [MemoryPackIgnore]
    public int Count => Cells.Count;

    [MemoryPackIgnore]
    public Expression this[string key]
    {
        get => Cells.TryGetValue(key, out var cell) ? cell : BlankValue.Instance;
        set => Cells[key] = value;
    }

    [MemoryPackIgnore]
    public IEnumerable<string> Keys => Cells.Keys;

    [MemoryPackIgnore]
    public IEnumerable<Expression> Values => Cells.Values;

    public IEnumerator<KeyValuePair<string, Expression>> GetEnumerator()
    {
        return Cells.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)Cells).GetEnumerator();
    }

    public bool ContainsKey(string key)
    {
        return Cells.ContainsKey(key);
    }

    public bool TryGetValue(string key, [MaybeNullWhen(false)] out Expression value)
    {
        return Cells.TryGetValue(key, out value);
    }
}
