using System.Collections.Concurrent;
using Danfma.MySheet.Expressions;
using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// Identifies the pure (criteria-free) aggregate a <see cref="RangeSnapshot"/> memoizes, so SUM/COUNT/…
/// of the SAME range repeated in thousands of formulas is folded once and served O(1) thereafter.
/// </summary>
internal enum AggregateKind
{
    Sum,
    Count,
    CountA,
    Max,
    Min,
    Average,
}

/// <summary>
/// Memoizes a criteria-free aggregate (SUM/COUNT/COUNTA/MAX/MIN/AVERAGE) whose only argument is a big
/// populated range: the SAME aggregate repeated across thousands of formulas is folded once (through the
/// function's OWN body, so the result is semantically identical) and served O(1) thereafter.
/// </summary>
internal static class RangeAggregate
{
    public static ComputedValue Memoize(
        Expression[] arguments,
        EvaluationContext context,
        AggregateKind kind,
        Func<ComputedValue> compute
    ) =>
        arguments is [Reference reference]
        && context.Workbook.TryGetRangeSnapshot(reference, context) is { } snapshot
            ? snapshot.Aggregate(kind, compute)
            : compute();
}

/// <summary>
/// A per-epoch admission slot for one range key. The cache admits a range only on its SECOND read: the first
/// read records a lightweight <c>Seen</c> marker (this entry with a null snapshot) and keeps the caller on its
/// linear path; the second read builds the <see cref="RangeSnapshot"/> and every read thereafter reuses it.
/// This spares the O(N) materialization for ranges a formula reads exactly ONCE per epoch (sliding windows,
/// single-shot bounded lookups, invalidate-heavy loops), which was the 2.6.2 production regression.
///
/// <para>The build is race-benign: two threads arriving on the second read may both build, but
/// <see cref="Interlocked.CompareExchange{T}(ref T,T,T)"/> keeps exactly one instance and the loser is
/// dropped — the same "build once, publish once" guarantee the rest of the workbook caches use.</para>
/// </summary>
internal sealed class RangeCacheEntry
{
    private RangeSnapshot? _snapshot;

    /// <summary>The built snapshot, or <c>null</c> while the entry is still a bare <c>Seen</c> marker.</summary>
    public RangeSnapshot? Snapshot => _snapshot;

    /// <summary>Returns the snapshot, building it on first demand (the range's second read). Concurrent callers
    /// may build in parallel but publish a single instance.</summary>
    public RangeSnapshot GetOrBuild(Reference range, EvaluationContext context)
    {
        var existing = _snapshot;
        if (existing is not null)
        {
            return existing;
        }

        var built = RangeSnapshot.Build(range, context);
        return Interlocked.CompareExchange(ref _snapshot, built, null) ?? built;
    }
}

/// <summary>Outcome of an exact-hash probe: the value's kind decides whether the hash can answer at all.</summary>
internal enum ExactMatchOutcome
{
    /// <summary>The hash answered: the value is present at <c>position</c> (its first occurrence).</summary>
    Found,

    /// <summary>The hash answered: the value is provably absent (safe to report #N/A).</summary>
    NotFound,

    /// <summary>The hash cannot answer (blank-equivalent, blank, error or NaN lookup); fall back to linear.</summary>
    Unsupported,
}

/// <summary>
/// Layer-2 per-epoch value cache for a single populated range (a whole column, a bounded rectangle, a
/// lookup key column). The <see cref="Values"/> snapshot is materialized ONCE via the Layer-1 structural
/// index + the memoized cell cache, and every derived accelerator (exact hash, sorted index, numeric
/// equality map, sorted-number view, aggregate memo) is built LAZILY on first demand and then shared by
/// every formula of the epoch — collapsing the O(F·N) re-enumeration the whole-column scenario pays.
///
/// <para>Semantics are preserved BIT FOR BIT: each derived view reproduces exactly what the current linear
/// scan of the consuming function computes (position numbering, error/blank skipping, tie-breaking). Where
/// a value's semantics do not fit an index cleanly (blank-equivalent lookups, wildcard/comparison criteria)
/// the consumer falls back to a LINEAR scan over <see cref="Values"/> — still cache-served, so the O(N)
/// re-read of cell values is gone even on the fallback path.</para>
///
/// <para>Not serialized; a plain internal class (never MemoryPack-deserialized), so field initializers and
/// <see cref="Lazy{T}"/> are safe here. Owned per-epoch by the <see cref="Workbook"/> and dropped by BOTH
/// <see cref="Workbook.InvalidateCache"/> and <see cref="Workbook.Recalculate"/> (a snapshot may carry
/// volatile-tainted values).</para>
/// </summary>
internal sealed class RangeSnapshot
{
    /// <summary>The cell values in the range's enumeration order — a 1-based position is <c>index + 1</c>,
    /// exactly the position the linear scans of MATCH/VLOOKUP/… report. For a closed rectangle this is a
    /// zero-copy arithmetic view straight over the dense store (see <see cref="RectangleValueList"/>) — no
    /// second copy of values that already live in <see cref="SheetValueStore"/>; for an open range it stays the
    /// materialized list (its enumeration order comes from the structural index, not arithmetic).</summary>
    public IReadOnlyList<ComputedValue> Values { get; }

    private readonly Lazy<ExactIndex> _exact;
    private readonly Lazy<SortedIndex> _sorted;
    private readonly Lazy<NumericEqualityMap> _numericEquality;
    private readonly Lazy<SortedNumbers> _sortedNumbers;
    private readonly ConcurrentDictionary<AggregateKind, ComputedValue> _aggregates = new();

    private RangeSnapshot(IReadOnlyList<ComputedValue> values)
    {
        Values = values;
        _exact = new Lazy<ExactIndex>(BuildExact);
        _sorted = new Lazy<SortedIndex>(BuildSorted);
        _numericEquality = new Lazy<NumericEqualityMap>(BuildNumericEquality);
        _sortedNumbers = new Lazy<SortedNumbers>(BuildSortedNumbers);
    }

    public int Count => Values.Count;

    /// <summary>Materializes the snapshot from a range reference, reusing the exact enumeration order and the
    /// memoized cell values the consuming functions already see (so the linear-fallback path is identical).</summary>
    public static RangeSnapshot Build(Reference range, EvaluationContext context)
    {
        switch (range)
        {
            case RangeReference rectangle:
                // A closed rectangle has known bounds AND arithmetic position<->(col,row): FORCE every covered
                // cell to be computed/memoized (block-copy-free presence probe when a column is already fully
                // present, else the per-cell numeric accessor), then hand back a zero-copy view over the dense
                // store — no second ComputedValue[] duplicating what SheetValueStore already retains for the
                // epoch. The view's element order is IDENTICAL to the old materialized array (column outer,
                // row inner).
                return new RangeSnapshot(BuildRectangleView(rectangle, context));

            case OpenRangeReference open:
                // Open ranges expose only their POPULATED cells (count unknown up front, order = the structural
                // index's, not arithmetic) — no cheap position<->(col,row) inverse exists (it would need an
                // auxiliary per-covered-column offset table read through the structural index), so the List
                // path is kept.
                var values = new List<ComputedValue>();
                foreach (var value in open.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                return new RangeSnapshot(values);

            default:
                return new RangeSnapshot(Array.Empty<ComputedValue>());
        }
    }

    /// <summary>
    /// Forces every cell of a closed rectangle to be computed and memoized in the dense store — a per-column
    /// presence probe (no copy) skips columns already fully present; otherwise the per-cell numeric accessor
    /// evaluates each miss on demand — then returns a zero-copy <see cref="RectangleValueList"/> view over the
    /// store. Equivalent BIT FOR BIT to the old materialized array: once this returns, every covered (col,row)
    /// is present in the store with exactly the value the old array would have captured, and — within the
    /// mainline epoch model (a cell present in the store never silently changes before
    /// <see cref="Workbook.InvalidateCache"/>/<see cref="Workbook.Recalculate"/> drop this very cache) — stays
    /// that way for the rest of the epoch.
    /// </summary>
    private static RectangleValueList BuildRectangleView(
        RangeReference rectangle,
        EvaluationContext context
    )
    {
        var start = CellAddress.Parse(rectangle.StartId);
        var end = CellAddress.Parse(rectangle.EndId);

        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        var rowCount = maxRow - minRow + 1;
        var columnCount = maxColumn - minColumn + 1;
        checked
        {
            _ = rowCount * columnCount; // preserve the old path's overflow guard
        }

        var workbook = context.Workbook;
        var store = workbook.DenseStore;
        var sheetName = rectangle.SheetName;
        var handle = store.HandleFor(sheetName);

        for (var column = minColumn; column <= maxColumn; column++)
        {
            // Fast path: the whole column's covered rows are already present -> nothing to force.
            if (store.IsColumnRangePresent(handle, column, minRow, maxRow))
            {
                continue;
            }

            // Fallback: per-cell numeric accessor. A hit is a lock-free store read; a miss routes through the
            // workbook for the cycle guard + on-demand evaluation + memoization (identical to the old path) —
            // the result is discarded here (only presence matters); the view reads it back on actual access.
            for (var row = minRow; row <= maxRow; row++)
            {
                if (!store.TryGetDense(handle, column, row, out _))
                {
                    workbook.GetCellValueDense(handle, sheetName, column, row);
                }
            }
        }

        return new RectangleValueList(
            workbook,
            store,
            handle,
            sheetName,
            minColumn,
            minRow,
            rowCount,
            columnCount
        );
    }

    /// <summary>
    /// Zero-copy <see cref="IReadOnlyList{T}"/> view of a closed rectangle's cells, addressed by the SAME
    /// column-major position <see cref="RangeSnapshot.Build"/> used to fill the old materialized array
    /// (<c>index = (column - minColumn) * rowCount + (row - minRow)</c>, inverted below). Every access reads
    /// straight through <see cref="SheetValueStore"/> (a lock-free, seqlock-protected read) instead of a
    /// retained copy — the store already holds every covered cell after <see cref="BuildRectangleView"/>
    /// forced it, so a hit is the overwhelmingly common case; a miss (a defensive fallback, never expected once
    /// <see cref="Build"/> has returned) still evaluates on demand exactly like the old per-cell path, so this
    /// is correct even if that invariant is ever violated.
    /// </summary>
    private sealed class RectangleValueList : IReadOnlyList<ComputedValue>
    {
        private readonly Workbook _workbook;
        private readonly SheetValueStore _store;
        private readonly int _handle;
        private readonly string _sheetName;
        private readonly int _minColumn;
        private readonly int _minRow;
        private readonly int _rowCount;

        public RectangleValueList(
            Workbook workbook,
            SheetValueStore store,
            int handle,
            string sheetName,
            int minColumn,
            int minRow,
            int rowCount,
            int columnCount
        )
        {
            _workbook = workbook;
            _store = store;
            _handle = handle;
            _sheetName = sheetName;
            _minColumn = minColumn;
            _minRow = minRow;
            _rowCount = rowCount;
            Count = rowCount * columnCount;
        }

        public int Count { get; }

        public ComputedValue this[int index]
        {
            get
            {
                if ((uint)index >= (uint)Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                var column = _minColumn + (index / _rowCount);
                var row = _minRow + (index % _rowCount);

                return _store.TryGetDense(_handle, column, row, out var value)
                    ? value
                    : _workbook.GetCellValueDense(_handle, _sheetName, column, row);
            }
        }

        public IEnumerator<ComputedValue> GetEnumerator()
        {
            for (var i = 0; i < Count; i++)
            {
                yield return this[i];
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
            GetEnumerator();
    }

    // === Aggregate memo (SUM/COUNT/COUNTA/MAX/MIN/AVERAGE of a pure range) ================================

    /// <summary>Returns the memoized result of a criteria-free aggregate over the range, computing it once
    /// (via the function's OWN body, so the result is semantically identical) and serving repeats O(1).</summary>
    public ComputedValue Aggregate(AggregateKind kind, Func<ComputedValue> compute)
    {
        if (_aggregates.TryGetValue(kind, out var cached))
        {
            return cached;
        }

        var computed = compute();
        _aggregates.TryAdd(kind, computed);
        return computed;
    }

    // === Exact hash (value → first 1-based position) =====================================================

    private sealed class ExactIndex
    {
        public readonly Dictionary<double, int> Numbers = new();
        public readonly Dictionary<string, int> Text = new(StringComparer.OrdinalIgnoreCase);
        public int TruePosition; // 0 = absent
        public int FalsePosition;
    }

    private ExactIndex BuildExact()
    {
        var index = new ExactIndex();

        for (var i = 0; i < Values.Count; i++)
        {
            var value = Values[i];
            var position = i + 1;

            switch (value.Kind)
            {
                case ComputedValueKind.Number:
                    value.TryGetNumber(out var number);
                    // NaN never equals anything under Excel's `=`, but Dictionary<double,int> treats NaN keys
                    // as equal — so it is excluded here and reported Unsupported on the probe side.
                    if (!double.IsNaN(number))
                    {
                        index.Numbers.TryAdd(number, position);
                    }

                    break;

                case ComputedValueKind.Text:
                    if (value.TryGetText(out var text))
                    {
                        index.Text.TryAdd(text, position);
                    }

                    break;

                case ComputedValueKind.Boolean:
                    value.TryGetBoolean(out var boolean);
                    if (boolean)
                    {
                        if (index.TruePosition == 0)
                        {
                            index.TruePosition = position;
                        }
                    }
                    else if (index.FalsePosition == 0)
                    {
                        index.FalsePosition = position;
                    }

                    break;

                // Blank/Error/Reference cells never participate in an exact match against a "safe" lookup.
            }
        }

        return index;
    }

    /// <summary>
    /// Exact-match probe reproducing <see cref="ValueCoercion.AreEqual"/>'s FIRST-position result. Only a
    /// lookup that cannot be blank-equivalent (a non-zero number, a non-empty text, TRUE) is answerable in
    /// O(1); a blank-equivalent value (0, "", FALSE), a blank/error or a NaN would need the intransitive
    /// blank rule, so it returns <see cref="ExactMatchOutcome.Unsupported"/> for a linear fallback.
    /// </summary>
    public ExactMatchOutcome TryExactPosition(in ComputedValue lookup, out int position)
    {
        position = 0;

        switch (lookup.Kind)
        {
            case ComputedValueKind.Number:
                lookup.TryGetNumber(out var number);
                if (number == 0 || double.IsNaN(number))
                {
                    return ExactMatchOutcome.Unsupported;
                }

                return _exact.Value.Numbers.TryGetValue(number, out position)
                    ? ExactMatchOutcome.Found
                    : ExactMatchOutcome.NotFound;

            case ComputedValueKind.Text:
                lookup.TryGetText(out var text);
                if (string.IsNullOrEmpty(text))
                {
                    return ExactMatchOutcome.Unsupported;
                }

                return _exact.Value.Text.TryGetValue(text, out position)
                    ? ExactMatchOutcome.Found
                    : ExactMatchOutcome.NotFound;

            case ComputedValueKind.Boolean:
                lookup.TryGetBoolean(out var boolean);
                if (!boolean)
                {
                    return ExactMatchOutcome.Unsupported; // FALSE is blank-equivalent
                }

                position = _exact.Value.TruePosition;
                return position > 0 ? ExactMatchOutcome.Found : ExactMatchOutcome.NotFound;

            default:
                return ExactMatchOutcome.Unsupported; // blank/error/reference lookup
        }
    }

    // === Sorted index (approximate MATCH / VLOOKUP-TRUE, any input order) =================================

    private sealed class SortedIndex
    {
        public required ComputedValue[] SortedValues;
        public required int[] PrefixMaxPosition; // max original position over SortedValues[0..k]
        public required int[] SuffixMaxPosition; // max original position over SortedValues[k..end]
    }

    private SortedIndex BuildSorted()
    {
        // Non-blank/non-error cells only — MATCH/VLOOKUP approximate scans skip both.
        var entries = new List<(ComputedValue Value, int Position)>(Values.Count);

        for (var i = 0; i < Values.Count; i++)
        {
            var value = Values[i];
            if (value.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
            {
                continue;
            }

            entries.Add((value, i + 1));
        }

        entries.Sort(static (a, b) => ValueCoercion.Compare(a.Value, b.Value));

        var count = entries.Count;
        var sortedValues = new ComputedValue[count];
        var prefix = new int[count];
        var suffix = new int[count];

        for (var i = 0; i < count; i++)
        {
            sortedValues[i] = entries[i].Value;
            prefix[i] = i == 0 ? entries[i].Position : Math.Max(prefix[i - 1], entries[i].Position);
        }

        for (var i = count - 1; i >= 0; i--)
        {
            suffix[i] =
                i == count - 1 ? entries[i].Position : Math.Max(suffix[i + 1], entries[i].Position);
        }

        return new SortedIndex
        {
            SortedValues = sortedValues,
            PrefixMaxPosition = prefix,
            SuffixMaxPosition = suffix,
        };
    }

    /// <summary>
    /// Reproduces MATCH type 1 / VLOOKUP-TRUE (largest value ≤ lookup, returning the LAST such position in
    /// scan order — Excel's "last of the duplicates"). Correct for ANY input order: the elements ≤ lookup
    /// form a prefix of the sorted view, and the answer is that prefix's maximum original position. Returns
    /// 0 (→ #N/A) when no populated value is ≤ lookup.
    /// </summary>
    public int ApproximateAscendingPosition(in ComputedValue lookup)
    {
        var index = _sorted.Value;
        var values = index.SortedValues;

        // upperBound: count of elements whose Compare(value, lookup) ≤ 0.
        var low = 0;
        var high = values.Length;
        while (low < high)
        {
            var mid = (low + high) >> 1;
            if (ValueCoercion.Compare(values[mid], lookup) <= 0)
            {
                low = mid + 1;
            }
            else
            {
                high = mid;
            }
        }

        return low == 0 ? 0 : index.PrefixMaxPosition[low - 1];
    }

    /// <summary>
    /// Reproduces MATCH type -1 (smallest value ≥ lookup, LAST such position in scan order). The elements
    /// ≥ lookup form a suffix of the sorted view; the answer is that suffix's maximum original position.
    /// Returns 0 (→ #N/A) when no populated value is ≥ lookup.
    /// </summary>
    public int ApproximateDescendingPosition(in ComputedValue lookup)
    {
        var index = _sorted.Value;
        var values = index.SortedValues;

        // lowerBound: first index whose Compare(value, lookup) ≥ 0.
        var low = 0;
        var high = values.Length;
        while (low < high)
        {
            var mid = (low + high) >> 1;
            if (ValueCoercion.Compare(values[mid], lookup) < 0)
            {
                low = mid + 1;
            }
            else
            {
                high = mid;
            }
        }

        return low == values.Length ? 0 : index.SuffixMaxPosition[low];
    }

    // === Numeric equality map (SUMIF/COUNTIF/AVERAGEIF with a numeric `=` criterion) ======================

    private sealed class NumericEqualityMap
    {
        public readonly Dictionary<double, (double Sum, int Count)> Map = new();
    }

    private NumericEqualityMap BuildNumericEquality()
    {
        var map = new NumericEqualityMap();

        foreach (var value in Values)
        {
            if (value.Kind != ComputedValueKind.Number)
            {
                continue;
            }

            value.TryGetNumber(out var number);
            if (double.IsNaN(number))
            {
                continue;
            }

            map.Map.TryGetValue(number, out var entry);
            map.Map[number] = (entry.Sum + number, entry.Count + 1);
        }

        return map;
    }

    /// <summary>The (sum-of-self, count) of the cells numerically equal to <paramref name="key"/> — serving
    /// COUNTIF, single-range SUMIF and single-range AVERAGEIF with a numeric equality criterion in O(1).</summary>
    public (double Sum, int Count) NumericEquality(double key)
    {
        if (double.IsNaN(key))
        {
            return (0, 0);
        }

        return _numericEquality.Value.Map.TryGetValue(key, out var entry) ? entry : (0, 0);
    }

    // === Sorted numbers (SMALL/LARGE/MEDIAN/PERCENTILE/QUARTILE via NumericAggregation.Fold semantics) ====

    private sealed class SortedNumbers
    {
        public required double[] Values;
        public Error? FirstError;
    }

    private SortedNumbers BuildSortedNumbers()
    {
        var numbers = new List<double>(Values.Count);
        Error? firstError = null;

        foreach (var value in Values)
        {
            if (value.TryGetError(out var error))
            {
                firstError ??= error;
            }
            else if (value.TryGetNumber(out var number))
            {
                numbers.Add(number);
            }

            // Referenced text/logicals/blanks are ignored, matching NumericAggregation.AddReferenced.
        }

        var array = numbers.ToArray();
        Array.Sort(array);

        return new SortedNumbers { Values = array, FirstError = firstError };
    }

    /// <summary>
    /// The ascending-sorted numeric values (SUM-style referenced gathering) and the first cell error in scan
    /// order, shared by the order statistics. When <paramref name="firstError"/> is set the caller returns it
    /// unchanged (exactly as <c>StatisticsMath.Collect</c> propagates the first error). The returned list is
    /// SHARED and read-only — callers must not mutate it.
    /// </summary>
    public IReadOnlyList<double> SortedNumericValues(out Error? firstError)
    {
        var data = _sortedNumbers.Value;
        firstError = data.FirstError;
        return data.Values;
    }

    /// <summary>
    /// Binary-search-derived tie counts for <paramref name="number"/> against the shared ascending-sorted
    /// numeric view: the count of populated numeric values strictly LESS than it, the count EQUAL to it,
    /// and the count strictly GREATER — the O(log n) replacement for RANK.EQ/RANK.AVG's linear
    /// equal/outranking scan (an O(n) scan per formula, O(n²) over a dragged-down column). As with
    /// <see cref="SortedNumericValues"/>, when <paramref name="firstError"/> is set the caller must
    /// propagate it unchanged and ignore the counts — exactly how <c>StatisticsMath.Collect</c> short-
    /// circuits the linear path on the first cell error in scan order.
    /// </summary>
    public (int CountLess, int CountEqual, int CountGreater) NumericRankCounts(
        double number,
        out Error? firstError
    )
    {
        var data = _sortedNumbers.Value;
        firstError = data.FirstError;

        var values = data.Values;
        var lower = LowerBound(values, number);
        var upper = UpperBound(values, number);

        return (lower, upper - lower, values.Length - upper);
    }

    // First index whose value is >= number — the count of values strictly less than it.
    private static int LowerBound(double[] values, double number)
    {
        var low = 0;
        var high = values.Length;

        while (low < high)
        {
            var mid = (low + high) >> 1;
            if (values[mid] < number)
            {
                low = mid + 1;
            }
            else
            {
                high = mid;
            }
        }

        return low;
    }

    // First index whose value is > number — the count of values less than or equal to it.
    private static int UpperBound(double[] values, double number)
    {
        var low = 0;
        var high = values.Length;

        while (low < high)
        {
            var mid = (low + high) >> 1;
            if (values[mid] <= number)
            {
                low = mid + 1;
            }
            else
            {
                high = mid;
            }
        }

        return low;
    }
}

public sealed partial class Workbook
{
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
    // Internal (not private): VLOOKUP/HLOOKUP's linear fallback probes this threshold against the table's own
    // row/column count BEFORE building the single-column/row RangeReference key, so a small table (the common
    // case) never pays that allocation just to have TryGetRangeSnapshot immediately reject it as too small.
    internal const int RangeCacheMinimumCells = 256;

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

    // A cheap estimate of a range's populated-cell count, used only to decide whether caching is worth it. An
    // open range sums the covered structural-index lists (ignoring row bounds — an over-estimate is harmless:
    // it only risks caching a range that turns out small). A bounded rectangle is population-AWARE when it can
    // afford to be (see the RangeReference case below); otherwise it falls back to its capped area.
    private int EstimatePopulatedCells(Reference range, EvaluationContext context)
    {
        switch (range)
        {
            case RangeReference rectangle:
                var bounds = rectangle.GetBounds();
                var area = (long)bounds.RowCount * bounds.ColumnCount;
                var cappedArea =
                    area >= RangeCacheMinimumCells ? RangeCacheMinimumCells : (int)area;

                if (area < RangeCacheMinimumCells)
                {
                    return cappedArea; // already below threshold on area alone; the index adds nothing here
                }

                // The rectangle clears the AREA threshold, but the area is BLIND to its actual population — a
                // tall, mostly-blank single column (e.g. 1000 rows × 1 column, 10 populated) would otherwise be
                // admitted on faith and materialize a mostly-blank snapshot. Consult the structural index for
                // an exact, population-aware count instead — but ONLY when the column map already exists (a
                // prior column-oriented open-range read on this sheet already paid its O(sheet) bucketize
                // cost): building it here, on a rectangle's read path, would trade this bounded waste (an
                // over-admitted small rectangle) for an unbounded one (an O(sheet) scan just to decide whether
                // a possibly-tiny rectangle is worth caching). Absent the map, the legacy capped-area estimate
                // is kept — the exact same behavior as before this fix.
                if (
                    Sheets.TryGetValue(rectangle.SheetName, out var rectSheet)
                    && rectSheet.PeekStructuralIndex() is { HasColumnBuckets: true } rectIndex
                )
                {
                    var rectTotal = 0;

                    for (var column = bounds.LeftColumn; column <= bounds.RightColumn; column++)
                    {
                        rectTotal += rectIndex.CountColumnRowsInRange(
                            column,
                            bounds.TopRow,
                            bounds.BottomRow,
                            RangeCacheMinimumCells - rectTotal
                        );

                        if (rectTotal >= RangeCacheMinimumCells)
                        {
                            break;
                        }
                    }

                    return rectTotal;
                }

                return cappedArea;

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
}
