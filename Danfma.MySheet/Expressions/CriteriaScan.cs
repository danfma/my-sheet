namespace Danfma.MySheet.Expressions;

/// <summary>
/// A forward-only positional cursor over a single range argument of a criteria aggregate (the SUMIFS
/// family). It NEVER materializes an intermediate list for a closed rectangle: when the shared per-epoch
/// snapshot is admitted it indexes the snapshot's <see cref="RangeSnapshot.Values"/> array directly
/// (zero-copy); otherwise it streams the memoized cells through <see cref="RangeValueSequence"/>'s
/// dispatch-free struct enumerator (dense positional reads, no allocation). Only the uncommon non-rectangle
/// shapes (open ranges/unions/scalars without a snapshot) fall back to the ordinary materialized list.
/// </summary>
internal struct PositionalRange
{
    // Exactly one backing is live: a list (snapshot array or the materialized fallback) OR the dense stream.
    private readonly IReadOnlyList<ComputedValue>? _list;
    private RangeValueSequence.Enumerator _cursor;
    private int _index;

    /// <summary>The cell count, known up front for every backing (array length, rectangle area, or list
    /// count) so the *IFS length validation never forces a materialization just to measure.</summary>
    public readonly int Count;

    private PositionalRange(IReadOnlyList<ComputedValue> list)
    {
        _list = list;
        _cursor = default;
        _index = 0;
        Count = list.Count;
    }

    private PositionalRange(RangeValueSequence.Enumerator cursor, int count)
    {
        _list = null;
        _cursor = cursor;
        _index = 0;
        Count = count;
    }

    /// <summary>
    /// Opens a cursor over one argument, preferring the cheapest backing: the admitted per-epoch snapshot
    /// (zero-copy over <see cref="RangeSnapshot.Values"/>) → the dense positional stream for a closed
    /// rectangle (no allocation) → a materialized list for open ranges/unions/scalars.
    /// </summary>
    public static PositionalRange Open(Expression argument, EvaluationContext context)
    {
        if (
            argument is Reference reference
            && context.Workbook.TryGetRangeSnapshot(reference, context) is { } snapshot
        )
        {
            return new PositionalRange(snapshot.Values);
        }

        if (argument is RangeReference rectangle)
        {
            return new PositionalRange(
                rectangle.ExpandComputedValues(context).GetEnumerator(),
                rectangle.RowCount * rectangle.ColumnCount
            );
        }

        return new PositionalRange(ArgumentFlattening.ExpandComputedValues(argument, context));
    }

    /// <summary>The next cell in position order (column-major, matching the materialized expansion exactly).
    /// Every parallel cursor is advanced once per position so they stay aligned.</summary>
    public ComputedValue Next()
    {
        if (_list is { } list)
        {
            return list[_index++];
        }

        _cursor.MoveNext();
        return _cursor.Current;
    }
}

/// <summary>
/// The shared position-by-position scan of the *IFS criteria aggregates (SUMIFS, COUNTIFS, AVERAGEIFS,
/// MAXIFS, MINIFS). It walks the criteria ranges (and the optional value range) as parallel forward
/// cursors — no intermediate <c>List&lt;ComputedValue&gt;</c> per range — after enforcing equal lengths up
/// front (<c>#VALUE!</c> on mismatch, like the pre-refactor code). Each <see cref="MoveNext"/> reports
/// whether every criterion matched at that position and, for the value families, hands back the
/// value-range cell so the caller's own accumulator (sum/count/average/extreme) stays in the function body.
/// </summary>
internal struct CriteriaScan
{
    private readonly PositionalRange[] _ranges;
    private readonly Criteria[] _criterias;
    private PositionalRange _valueRange;
    private readonly bool _hasValue;
    private readonly int _length;
    private int _index;

    private CriteriaScan(
        PositionalRange[] ranges,
        Criteria[] criterias,
        PositionalRange valueRange,
        bool hasValue,
        int length
    )
    {
        _ranges = ranges;
        _criterias = criterias;
        _valueRange = valueRange;
        _hasValue = hasValue;
        _length = length;
        _index = 0;
    }

    /// <summary>
    /// Builds a scan whose length is the VALUE range's cell count — the SUMIFS/AVERAGEIFS/MAXIFS/MINIFS
    /// shape where <c>arguments[0]</c> is the value range and the (criteria_range, criteria) pairs follow.
    /// Returns <c>#REF!</c> for a missing sheet and <c>#VALUE!</c> when a criteria range's length differs.
    /// </summary>
    public static Error? CreateWithValue(
        Expression[] arguments,
        EvaluationContext context,
        out CriteriaScan scan
    )
    {
        scan = default;

        if (ReferenceGuard.MissingSheet(arguments, context) is { } missing)
        {
            return missing;
        }

        var valueRange = PositionalRange.Open(arguments[0], context);
        var length = valueRange.Count;
        var pairCount = (arguments.Length - 1) / 2;
        var ranges = new PositionalRange[pairCount];
        var criterias = new Criteria[pairCount];

        for (var p = 0; p < pairCount; p++)
        {
            var range = PositionalRange.Open(arguments[1 + (p * 2)], context);

            if (range.Count != length)
            {
                return Error.Value;
            }

            ranges[p] = range;
            criterias[p] = Criteria.Parse(arguments[2 + (p * 2)].Evaluate(context));
        }

        scan = new CriteriaScan(ranges, criterias, valueRange, hasValue: true, length);

        return null;
    }

    /// <summary>
    /// Builds a value-less scan — the COUNTIFS shape where the (criteria_range, criteria) pairs start at
    /// <c>arguments[0]</c> and the length is the first criteria range's cell count. Returns <c>#REF!</c>
    /// for a missing sheet and <c>#VALUE!</c> when a later criteria range's length differs.
    /// </summary>
    public static Error? CreateCountOnly(
        Expression[] arguments,
        EvaluationContext context,
        out CriteriaScan scan
    )
    {
        scan = default;

        if (ReferenceGuard.MissingSheet(arguments, context) is { } missing)
        {
            return missing;
        }

        var pairCount = arguments.Length / 2;
        var ranges = new PositionalRange[pairCount];
        var criterias = new Criteria[pairCount];
        var length = 0;

        for (var p = 0; p < pairCount; p++)
        {
            var range = PositionalRange.Open(arguments[p * 2], context);

            if (p == 0)
            {
                length = range.Count;
            }
            else if (range.Count != length)
            {
                return Error.Value;
            }

            ranges[p] = range;
            criterias[p] = Criteria.Parse(arguments[(p * 2) + 1].Evaluate(context));
        }

        scan = new CriteriaScan(ranges, criterias, default, hasValue: false, length);

        return null;
    }

    /// <summary>
    /// Advances every cursor by one position (keeping them aligned) and returns <c>false</c> once the range
    /// is exhausted. On <c>true</c>, <paramref name="matched"/> tells whether every criterion matched at
    /// this position; <paramref name="value"/> is the value-range cell (default for COUNTIFS). The criteria
    /// checks short-circuit — the cursors still all advance, but the expensive <see cref="Criteria.Matches"/>
    /// is skipped once a mismatch is known, exactly like the old <c>AllMatch</c>.
    /// </summary>
    public bool MoveNext(out bool matched, out ComputedValue value)
    {
        if (_index >= _length)
        {
            matched = false;
            value = default;

            return false;
        }

        matched = true;
        var ranges = _ranges;
        var criterias = _criterias;

        for (var j = 0; j < ranges.Length; j++)
        {
            var cell = ranges[j].Next();

            if (matched && !criterias[j].Matches(cell))
            {
                matched = false;
            }
        }

        value = _hasValue ? _valueRange.Next() : default;
        _index++;

        return true;
    }
}
