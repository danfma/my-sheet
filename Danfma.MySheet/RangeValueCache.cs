using System.Collections.Concurrent;
using Danfma.MySheet.Expressions;

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
    /// <summary>The materialized cell values in the range's enumeration order — a 1-based position is
    /// <c>index + 1</c>, exactly the position the linear scans of MATCH/VLOOKUP/… report.</summary>
    public ComputedValue[] Values { get; }

    private readonly Lazy<ExactIndex> _exact;
    private readonly Lazy<SortedIndex> _sorted;
    private readonly Lazy<NumericEqualityMap> _numericEquality;
    private readonly Lazy<SortedNumbers> _sortedNumbers;
    private readonly ConcurrentDictionary<AggregateKind, ComputedValue> _aggregates = new();

    private RangeSnapshot(ComputedValue[] values)
    {
        Values = values;
        _exact = new Lazy<ExactIndex>(BuildExact);
        _sorted = new Lazy<SortedIndex>(BuildSorted);
        _numericEquality = new Lazy<NumericEqualityMap>(BuildNumericEquality);
        _sortedNumbers = new Lazy<SortedNumbers>(BuildSortedNumbers);
    }

    public int Count => Values.Length;

    /// <summary>Materializes the snapshot from a range reference, reusing the exact enumeration order and the
    /// memoized cell values the consuming functions already see (so the linear-fallback path is identical).</summary>
    public static RangeSnapshot Build(Reference range, EvaluationContext context)
    {
        switch (range)
        {
            case RangeReference rectangle:
                // A closed rectangle has known bounds: materialize straight into a PRE-SIZED array (no List
                // doubling + ToArray recopy), filled per column in column-major order — the block-copy fast path
                // when the column is fully present, else the per-cell numeric accessor (which evaluates misses on
                // demand). The element order is IDENTICAL to the old ExpandComputedValues path (column outer,
                // row inner).
                return new RangeSnapshot(BuildRectangle(rectangle, context));

            case OpenRangeReference open:
                // Open ranges expose only their POPULATED cells (count unknown up front, order = the structural
                // index's) — no pre-sizing/block-copy applies (Phase 3 territory), so the List path is kept.
                var values = new List<ComputedValue>();
                foreach (var value in open.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                return new RangeSnapshot(values.ToArray());

            default:
                return new RangeSnapshot(Array.Empty<ComputedValue>());
        }
    }

    /// <summary>
    /// Materializes a closed rectangle into a pre-sized column-major <see cref="ComputedValue"/>[] (one entry
    /// per cell, blanks included). Each column is filled by a single seqlock-verified block copy of its pages
    /// when every covered slot is present; otherwise the per-cell numeric accessor fills it, taking the hit path
    /// (a lock-free store read) or the on-demand evaluation of a miss — the exact semantics, and the exact
    /// column-major order, of the enumerator this replaced.
    /// </summary>
    private static ComputedValue[] BuildRectangle(
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
        var result = new ComputedValue[checked(rowCount * columnCount)];

        var workbook = context.Workbook;
        var store = workbook.DenseStore;
        var sheetName = rectangle.SheetName;
        var handle = store.HandleFor(sheetName);

        for (var column = minColumn; column <= maxColumn; column++)
        {
            var destBase = (column - minColumn) * rowCount;

            // Fast path: the whole column's covered rows are present -> block-copy its pages into dest.
            if (store.TryBlockCopyColumn(handle, column, minRow, maxRow, result, destBase))
            {
                continue;
            }

            // Fallback: per-cell numeric accessor. A hit is a lock-free store read; a miss routes through the
            // workbook for the cycle guard + on-demand evaluation + memoization (identical to the old path).
            for (var row = minRow; row <= maxRow; row++)
            {
                if (!store.TryGetDense(handle, column, row, out var value))
                {
                    value = workbook.GetCellValueDense(handle, sheetName, column, row);
                }

                result[destBase + (row - minRow)] = value;
            }
        }

        return result;
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

        for (var i = 0; i < Values.Length; i++)
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
        var entries = new List<(ComputedValue Value, int Position)>(Values.Length);

        for (var i = 0; i < Values.Length; i++)
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
        var numbers = new List<double>(Values.Length);
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
}
