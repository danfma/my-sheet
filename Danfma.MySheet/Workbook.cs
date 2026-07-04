using System.Buffers.Binary;
using System.Collections;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
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

    // === Volatile-function epoch model (F1) ===============================================================
    // A volatile cell (NOW/TODAY/RAND/RANDBETWEEN, directly or transitively) is cached WITHIN an epoch and its
    // address is recorded as tainted inside the value store. Recalculate() drops just those cells and re-samples
    // the clock (epoch++); InvalidateCache() drops everything. This keeps NOW()/RAND() coherent within a pass
    // (sampled once) while staying cheap to refresh, all without a dependency graph.

    // "The cell currently being evaluated on this thread touched a volatile." Same thread-local save/reset/
    // propagate pattern as _evaluating: GetCellValue zeroes it before a cell, reads it after (to mark the
    // cell), and ORs it back into the caller's value so volatility propagates up the evaluation stack.
    [ThreadStatic]
    private static bool _volatileTouched;

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

    // === Layer-1 structural index (whole-column scale) ==================================================
    // The per-sheet "which cells exist in this column/row" index now lives ON the Sheet (see
    // Sheet.GetStructuralIndex): it is write-maintained by the SetCell/Remove choke point and lifetime-scoped,
    // so InvalidateCache() no longer drops it (only cell edits change structure, and those maintain it in
    // place). The Workbook keeps no structural state of its own anymore.

    // === Layer-2 range value cache (whole-column scale) =================================================
    // Per-epoch snapshot (materialized ComputedValue[] + lazy derived accelerators) of a populated range,
    // keyed by the range reference record (OpenRangeReference/RangeReference have value equality). The
    // snapshot re-reads cell VALUES once and every consuming formula of the epoch then serves its lookup /
    // criterion / aggregate O(1)/O(log n) instead of re-scanning N cells. Values can be volatile-tainted, so
    // this is dropped by BOTH InvalidateCache() AND Recalculate() (unlike the structural index). Lazily
    // created race-free via the RangeCache accessor — never `= new()` on the field (MemoryPack bypasses it).
    [MemoryPackIgnore]
    private ConcurrentDictionary<Reference, RangeCacheEntry>? _rangeCache;

    // A range with fewer than this many populated cells is not cached: a linear scan already wins there and
    // caching every tiny range would flood the dictionary. Measured against the whole-column benchmark.
    private const int RangeCacheMinimumCells = 256;

    // Defensive cap on lightweight "Seen" markers. A workload of single-use ranges that clear the threshold
    // (e.g. 10k distinct sliding windows) only ever leaves markers behind — never a built snapshot — so the
    // marker set could grow unbounded within an epoch. Each marker is tiny (a key + a null-snapshot slot) and
    // drops on InvalidateCache/Recalculate, but past this count new ranges simply stay on the linear path
    // rather than pay even the marker. 64k markers is far above any legitimate reuse-heavy working set.
    private const int RangeCacheMarkerCap = 65_536;

    // Approximate count of live entries (markers + built). Interlocked, reset when the cache is dropped; used
    // only to enforce RangeCacheMarkerCap cheaply (ConcurrentDictionary.Count would lock every bucket).
    private int _rangeCacheEntryCount;

    private ConcurrentDictionary<Reference, RangeCacheEntry> RangeCache
    {
        get
        {
            var existing = _rangeCache;
            if (existing is not null)
            {
                return existing;
            }

            var created = new ConcurrentDictionary<Reference, RangeCacheEntry>();
            return Interlocked.CompareExchange(ref _rangeCache, created, null) ?? created;
        }
    }

    // Test-only bypass: forces every consumer down the pre-cache linear path, so the equivalence harness can
    // capture the "no cache" expectation and diff it against the cache-served result. Not serialized.
    [MemoryPackIgnore]
    internal bool RangeCacheDisabled { get; set; }

    /// <summary>
    /// The shared per-epoch <see cref="RangeSnapshot"/> for a populated range, or <c>null</c> when the range
    /// is not cacheable (not a rectangle/open range, below the size threshold, the cache is disabled, or this
    /// is the range's FIRST read this epoch) — in which case the caller keeps its existing linear path.
    ///
    /// <para><b>Second-use admission (Phase 4).</b> Materializing a snapshot pays off only when a range is read
    /// more than once per epoch. A range read exactly once (a sliding window, a one-shot bounded lookup) would
    /// pay an O(N) build it never reuses — the 2.6.2 regression. So the first read only records a lightweight
    /// marker and returns <c>null</c> (the caller's linear path is used); the snapshot is built on the SECOND
    /// read and reused by every read after. Ranges below the threshold are never even marked.</para>
    /// </summary>
    internal RangeSnapshot? TryGetRangeSnapshot(Reference range, EvaluationContext context)
    {
        if (RangeCacheDisabled || range is not (RangeReference or OpenRangeReference))
        {
            return null;
        }

        var cache = RangeCache;

        // Already seen this epoch: the second (or later) read builds the snapshot (once) and reuses it.
        if (cache.TryGetValue(range, out var entry))
        {
            return entry.GetOrBuild(range, context);
        }

        // First read: only worth remembering if it clears the size threshold and we are under the marker cap.
        if (
            Volatile.Read(ref _rangeCacheEntryCount) >= RangeCacheMarkerCap
            || EstimatePopulatedCells(range, context) < RangeCacheMinimumCells
        )
        {
            return null;
        }

        // Record a bare "Seen" marker and keep this first read on the linear path.
        if (cache.TryAdd(range, new RangeCacheEntry()))
        {
            Interlocked.Increment(ref _rangeCacheEntryCount);
        }

        return null;
    }

    // A cheap UPPER BOUND on a range's populated-cell count, used only to decide whether caching is worth it.
    // A bounded rectangle uses its area; an open range sums the covered structural-index lists (ignoring row
    // bounds — an over-estimate is harmless: it only risks caching a range that turns out small).
    private int EstimatePopulatedCells(Reference range, EvaluationContext context)
    {
        switch (range)
        {
            case RangeReference rectangle:
                var area = (long)rectangle.RowCount * rectangle.ColumnCount;
                return area >= RangeCacheMinimumCells ? RangeCacheMinimumCells : (int)area;

            case OpenRangeReference open:
                if (!Sheets.TryGetValue(open.SheetName, out var sheet))
                {
                    return 0;
                }

                // The structural index is lifetime-scoped and write-maintained, so consulting it is cheap: it
                // is built once per sheet life (never per epoch) and every read after is O(covered lists). The
                // count it gives is exact, so a small open range read repeatedly is correctly kept off the
                // value cache (below the threshold) instead of admitted on faith.
                var index = sheet.GetStructuralIndex();
                var total = 0;

                // Whole-row shape (both column limits open, row axis bounded): sum the covered rows.
                if (open is { ColMin: null, ColMax: null, RowMin: { } rowMin, RowMax: { } rowMax })
                {
                    for (var row = rowMin; row <= rowMax && total < RangeCacheMinimumCells; row++)
                    {
                        total += index.RowLength(row);
                    }

                    return total;
                }

                // Column-driven shapes: sum the covered columns' lengths (an upper bound over row bounds).
                foreach (var column in index.ColumnKeys)
                {
                    if (
                        (open.ColMin is not { } min || column >= min)
                        && (open.ColMax is not { } max || column <= max)
                    )
                    {
                        total += index.ColumnLength(column);
                        if (total >= RangeCacheMinimumCells)
                        {
                            break;
                        }
                    }
                }

                return total;

            default:
                return 0;
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
        // Drop from the store exactly the cells that touched a volatile this epoch; they recompute lazily on
        // the next read while every stable cell stays memoized.
        _valueStore?.DropTainted();

        // The value snapshots may hold volatile-tainted values, so they are dropped on a volatile refresh too
        // (the structural index — pure structure — survives).
        _rangeCache?.Clear();
        Volatile.Write(ref _rangeCacheEntryCount, 0);

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

    // === Save container format ===========================================================================
    // A cold, uncompressed Save writes the RAW MemoryPack of the model, byte-identical to every prior version
    // (a permanent regression contract). Any other combination wraps that model in a self-describing container
    // whose fixed 9-byte header is shared by every version:
    //   magic "MSWM" (4) | version (1) | modelLength (int32 LE, 4) | body
    // `version` selects the body encoding (the header layout and offsets never change — new warm-start files
    // stay v1, so their tests are unaffected):
    //   v1 (uncompressed): body = model bytes | value-block bytes            (warm-start, unchanged)
    //   v2 (Brotli):       body = Brotli(model bytes | value-block bytes)    (cold OR warm, compressed)
    // In every version `modelLength` is the UNCOMPRESSED model length, used to slice the (decompressed) body
    // into model vs. values. The value block is the MemoryPack of a List<CachedCellValue> surrogate (empty for
    // a cold compressed save). Load sniffs the 4-byte magic: a match is a container, anything else is a raw
    // (legacy or cold) model — the raw MemoryPack object header is a small member count (Workbook = 0x02),
    // never 'M' (0x4D), so the two are unambiguous.
    private static ReadOnlySpan<byte> ContainerMagic => "MSWM"u8;
    private const byte ContainerVersionUncompressed = 1;
    private const byte ContainerVersionBrotli = 2;
    private const int ContainerHeaderLength = 4 + 1 + 4; // magic + version + modelLength

    /// <summary>
    /// Serializes the workbook to a file (MemoryPack). The cell cache and registered custom functions are
    /// not persisted — they are rebuilt/re-registered after loading. Byte-identical across versions.
    /// </summary>
    public void Save(string path) => File.WriteAllBytes(path, MemoryPackSerializer.Serialize(this));

    /// <summary>
    /// Serializes the workbook to a file, honoring <paramref name="options"/>. With
    /// <see cref="WorkbookSaveOptions.IncludeComputedValues"/> the memoized values travel with the model in a
    /// container so a later <see cref="Load(string)"/> starts warm (no recomputation); volatile and
    /// reference-typed cache entries are never persisted. With
    /// <see cref="WorkbookSaveOptions.Compression"/> set to <see cref="WorkbookCompression.Brotli"/> the payload
    /// is Brotli-compressed inside the container. When both are at their defaults
    /// (<see cref="WorkbookCompression.None"/>, no values) the file is byte-identical to
    /// <see cref="Save(string)"/>. Either overload's output is transparently read back by
    /// <see cref="Load(string)"/>.
    /// </summary>
    public void Save(string path, WorkbookSaveOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        File.WriteAllBytes(path, SerializeToBytes(options));
    }

    /// <inheritdoc cref="Save(string)"/>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        await using var stream = File.Create(path);
        await MemoryPackSerializer.SerializeAsync(
            stream,
            this,
            cancellationToken: cancellationToken
        );
    }

    /// <inheritdoc cref="Save(string, WorkbookSaveOptions)"/>
    public async Task SaveAsync(
        string path,
        WorkbookSaveOptions options,
        CancellationToken cancellationToken = default
    )
    {
        ArgumentNullException.ThrowIfNull(options);
        await File.WriteAllBytesAsync(path, SerializeToBytes(options), cancellationToken);
    }

    // The bytes a Save writes. A cold, uncompressed save is the raw model, byte-identical to Save(path). Every
    // other combination is a container: warm (uncompressed) stays v1; anything compressed is v2 (Brotli over
    // model||values as ONE stream — a single stream compresses better than two separate blocks).
    private byte[] SerializeToBytes(WorkbookSaveOptions options)
    {
        var model = MemoryPackSerializer.Serialize(this);
        var compress = options.Compression == WorkbookCompression.Brotli;

        // Cold + uncompressed → the historical raw format (permanent byte-identity contract).
        if (!options.IncludeComputedValues && !compress)
        {
            return model;
        }

        // A cold compressed save carries no values; a warm save snapshots the cache.
        var values = MemoryPackSerializer.Serialize(
            options.IncludeComputedValues ? SnapshotComputedValues() : new List<CachedCellValue>()
        );

        return compress
            ? BuildContainer(ContainerVersionBrotli, model, BrotliCompress(model, values))
            : BuildContainer(ContainerVersionUncompressed, model, Concat(model, values));
    }

    // Prepends the fixed 9-byte header to an already-encoded body. `modelLength` is always the UNCOMPRESSED
    // model length so the reader can slice model vs. values after (optionally) decompressing the body.
    private static byte[] BuildContainer(byte version, byte[] model, byte[] body)
    {
        var buffer = new byte[ContainerHeaderLength + body.Length];
        var span = buffer.AsSpan();

        ContainerMagic.CopyTo(span);
        span[4] = version;
        BinaryPrimitives.WriteInt32LittleEndian(span.Slice(5, 4), model.Length);
        body.CopyTo(span.Slice(ContainerHeaderLength));

        return buffer;
    }

    private static byte[] Concat(byte[] first, byte[] second)
    {
        var buffer = new byte[first.Length + second.Length];
        first.CopyTo(buffer.AsSpan());
        second.CopyTo(buffer.AsSpan(first.Length));
        return buffer;
    }

    // Compresses model||values as a single Brotli stream (Optimal). One stream compresses better than two
    // independently-compressed blocks because the coder shares context across the whole payload.
    private static byte[] BrotliCompress(byte[] model, byte[] values)
    {
        using var output = new MemoryStream();
        using (var brotli = new BrotliStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            brotli.Write(model);
            brotli.Write(values);
        }

        return output.ToArray();
    }

    // Snapshots the memoized cache into surrogates, EXCLUDING volatile-tainted entries (they must re-sample on
    // the next read after loading) and reference-typed values (unrepresentable — their cells recompute).
    private List<CachedCellValue> SnapshotComputedValues()
    {
        var list = new List<CachedCellValue>();

        if (_valueStore is not { } store)
        {
            return list;
        }

        // The store already excludes volatile-tainted cells; the surrogate factory drops the unrepresentable
        // Reference kind. Present blank cells ARE carried (an explicitly-empty cached cell round-trips).
        foreach (var (sheetName, id, value) in store.EnumerateNonTainted())
        {
            if (CachedCellValue.TryFrom(sheetName, id, value) is { } surrogate)
            {
                list.Add(surrogate);
            }
        }

        return list;
    }

    // Repopulates the memoized store from a warm value block, reusing the same lazy store creation path as
    // GetCellValue so the field survives MemoryPack's field-initializer bypass and stays consistent.
    private void LoadComputedValues(List<CachedCellValue> values)
    {
        if (values.Count == 0)
        {
            return;
        }

        var store = ValueStore;

        foreach (var entry in values)
        {
            store.LoadEntry(entry.SheetName, entry.CellId, entry.ToComputedValue());
        }
    }

    /// <summary>Loads a workbook from a file written by <see cref="Save(string)"/> and returns the new
    /// instance. Warm-start containers repopulate the value cache; raw (cold/legacy) files load unchanged.</summary>
    public static Workbook Load(string path) => Deserialize(File.ReadAllBytes(path), path);

    /// <inheritdoc cref="Load(string)"/>
    public static async Task<Workbook> LoadAsync(
        string path,
        CancellationToken cancellationToken = default
    ) => Deserialize(await File.ReadAllBytesAsync(path, cancellationToken), path);

    // Sniffs the 4-byte magic: a container repopulates the warm cache; anything else is a raw model.
    private static Workbook Deserialize(byte[] bytes, string path)
    {
        if (
            bytes.Length >= ContainerHeaderLength
            && bytes.AsSpan(0, ContainerMagic.Length).SequenceEqual(ContainerMagic)
        )
        {
            return DeserializeContainer(bytes, path);
        }

        return MemoryPackSerializer.Deserialize<Workbook>(bytes)
            ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");
    }

    private static Workbook DeserializeContainer(byte[] bytes, string path)
    {
        var version = bytes[4];
        var modelLength = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(5, 4));

        if (modelLength < 0)
        {
            throw new InvalidDataException($"'{path}' is a corrupt MySheet container.");
        }

        // Resolve the (decompressed) body once, then slice model vs. values by modelLength. For v1 the body is
        // the bytes after the header verbatim; for v2 it is the Brotli-decompressed payload.
        var body = version switch
        {
            ContainerVersionUncompressed => bytes.AsMemory(ContainerHeaderLength),
            ContainerVersionBrotli => BrotliDecompress(bytes.AsSpan(ContainerHeaderLength), path),
            _ => throw new InvalidDataException(
                $"'{path}' is a MySheet container of unsupported version {version}."
            ),
        };

        if (modelLength > body.Length)
        {
            throw new InvalidDataException($"'{path}' is a corrupt MySheet container.");
        }

        var workbook =
            MemoryPackSerializer.Deserialize<Workbook>(body.Span.Slice(0, modelLength))
            ?? throw new InvalidDataException($"'{path}' did not contain a workbook.");

        var values =
            MemoryPackSerializer.Deserialize<List<CachedCellValue>>(body.Span.Slice(modelLength))
            ?? new List<CachedCellValue>();

        workbook.LoadComputedValues(values);

        return workbook;
    }

    private static byte[] BrotliDecompress(ReadOnlySpan<byte> compressed, string path)
    {
        try
        {
            using var input = new MemoryStream(compressed.ToArray());
            using var brotli = new BrotliStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            brotli.CopyTo(output);
            return output.ToArray();
        }
        catch (Exception inner)
            when (inner is InvalidDataException or IOException or InvalidOperationException)
        {
            throw new InvalidDataException(
                $"'{path}' is a corrupt MySheet container (Brotli payload could not be decompressed).",
                inner
            );
        }
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
            BinaryOperation binary => HasUnqualifiedReference(binary.Left)
                || HasUnqualifiedReference(binary.Right),
            UnaryOperation unary => HasUnqualifiedReference(unary.Operand),
            Function function => FormulaWriter
                .Call(function)
                .Arguments.Any(HasUnqualifiedReference),
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

    // The cell store is a PRIVATE field carrying the serialized member, at the exact declaration position the
    // public `Cells` property held before — MemoryPack orders members by declaration, so this keeps the wire
    // schema (member #3, a Dictionary<string, Expression>) byte-identical (proven by the pre-namespaces
    // fixture). The initializer runs only for a fresh `new Sheet`; MemoryPack bypasses field initializers on
    // deserialize, but every serialized file carries this member, so the field is never left null.
    [MemoryPackInclude]
    private Dictionary<string, Expression> _cells = new();

    /// <summary>
    /// Read-only view of the sheet's populated cells. Mutation goes through the write choke point
    /// (<see cref="SetCell"/>, reached via the indexer <c>set</c>) and <see cref="Remove"/> — the two, and only,
    /// paths that change the cell store.
    /// </summary>
    [MemoryPackIgnore]
    public IReadOnlyDictionary<string, Expression> Cells => _cells;

    // The write-maintained structural index (whole-column scale), owned per-Sheet because the write choke
    // point lives here. Runtime-only: [MemoryPackIgnore] so it never touches the wire schema, and lazily
    // created race-free via GetStructuralIndex (never `= new()` on the field — MemoryPack bypasses field
    // initializers on deserialize, so a loaded sheet starts with a null index and rebuilds it once on its
    // first open-range read). Built lazily, then kept up to date by SetCell/Remove; it SURVIVES
    // Workbook.InvalidateCache (structure is orthogonal to the value caches).
    [MemoryPackIgnore]
    private SheetStructuralIndex? _structuralIndex;

    /// <summary>
    /// The sheet's structural index, created (empty, unbuilt) on first access and reused for the sheet's life.
    /// The first read of a map on it triggers its one O(N) build; every <see cref="SetCell"/>/<see cref="Remove"/>
    /// keeps it current thereafter. Internal — part of the evaluation contract, not host API.
    /// </summary>
    internal SheetStructuralIndex GetStructuralIndex()
    {
        var existing = _structuralIndex;
        if (existing is not null)
        {
            return existing;
        }

        var created = new SheetStructuralIndex(this);
        return Interlocked.CompareExchange(ref _structuralIndex, created, null) ?? created;
    }

    /// <summary>The already-created structural index, or <c>null</c> when none has been created yet — a
    /// side-effect-free peek (never creates or builds). Test/introspection only.</summary>
    internal SheetStructuralIndex? PeekStructuralIndex() => _structuralIndex;

    [MemoryPackIgnore]
    public int Count => _cells.Count;

    [MemoryPackIgnore]
    public Expression this[string key]
    {
        get => _cells.TryGetValue(key, out var cell) ? cell : BlankValue.Instance;
        set => SetCell(key, value);
    }

    /// <summary>
    /// The single write path into the cell store: the public indexer <c>set</c> delegates here, so every cell
    /// insert or overwrite funnels through one method. It maintains the write-maintained structural index in
    /// place — an insert of a NEW id is recorded incrementally (an in-order append stays O(1); an out-of-order
    /// insert dirties only its bucket), while OVERWRITING an existing id leaves the index untouched (the id is
    /// already indexed). The index is only touched once it exists: before the sheet's first open-range read
    /// there is nothing to maintain (the eventual lazy build captures every current cell). It is the documented
    /// attach point for the future reverse dependency graph too. Like the indexer it replaced, it does NOT
    /// invalidate memoized values; the host calls <see cref="Workbook.InvalidateCache"/> after editing
    /// (see <see cref="Remove"/> for the symmetric delete).
    /// </summary>
    internal void SetCell(string id, Expression expr)
    {
        // One hash lookup: get (or add) the slot and learn whether the id already existed, so an overwrite
        // skips index maintenance while a genuinely new cell updates it.
        ref var slot = ref CollectionsMarshal.GetValueRefOrAddDefault(_cells, id, out var existed);
        slot = expr;

        if (!existed && _structuralIndex is { } index && CellAddress.TryGetColumnRow(id, out var column, out var row))
        {
            index.OnCellAdded(column, row, id);
        }
    }

    /// <summary>
    /// Removes a cell from the sheet, returning <c>true</c> when it existed (and <c>false</c> for a no-op).
    /// The second mutation path alongside the <see cref="SetCell"/> write choke point: it removes the id from
    /// the structural index (in place, O(bucket)) when the index exists, and — like a write — it does NOT
    /// invalidate memoized values: removing a cell changes the result of every formula that read it, so the
    /// host calls <see cref="Workbook.InvalidateCache"/> afterwards for the change to be observed, exactly as
    /// after a write. It is the delete-side attach point for the same future dependency graph as
    /// <see cref="SetCell"/>.
    /// </summary>
    public bool Remove(string id)
    {
        if (!_cells.Remove(id))
        {
            return false;
        }

        if (_structuralIndex is { } index && CellAddress.TryGetColumnRow(id, out var column, out var row))
        {
            index.OnCellRemoved(column, row, id);
        }

        return true;
    }

    [MemoryPackIgnore]
    public IEnumerable<string> Keys => _cells.Keys;

    [MemoryPackIgnore]
    public IEnumerable<Expression> Values => _cells.Values;

    public IEnumerator<KeyValuePair<string, Expression>> GetEnumerator()
    {
        return _cells.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)_cells).GetEnumerator();
    }

    public bool ContainsKey(string key)
    {
        return _cells.ContainsKey(key);
    }

    public bool TryGetValue(string key, [MaybeNullWhen(false)] out Expression value)
    {
        return _cells.TryGetValue(key, out value);
    }
}
