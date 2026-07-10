namespace Danfma.MySheet.Expressions;

/// <summary>
/// A forward-only, allocation-conscious cursor over the memoized values of a SINGLE range argument — the
/// non-admitted (small-range) fallback of the lookup/count scans (MATCH, XLOOKUP, COUNTIF). It NEVER
/// materializes an intermediate <see cref="List{T}"/>: an admitted per-epoch snapshot is indexed directly
/// (zero-copy over <see cref="RangeSnapshot.Values"/>); a closed rectangle streams through
/// <see cref="RangeValueSequence"/>'s dispatch-free struct enumerator (zero allocation); only the unbounded
/// shapes (open ranges, unions, a reference-valued scalar) fall back to a single boxed iterator — one small
/// state machine, versus the whole materialized vector the scan used to build per evaluation.
///
/// <para>The value order is identical, cell for cell, to
/// <see cref="ArgumentFlattening.ExpandComputedValues(Expression, EvaluationContext)"/> — every backing walks
/// the same enumeration — so a linear scan over the cursor stays bit-for-bit equivalent to the old scan over
/// the materialized list.</para>
/// </summary>
internal struct RangeValueCursor
{
    // Exactly one backing is live: an indexed list (admitted snapshot's array), the dense struct stream
    // (closed rectangle) or a boxed iterator (open range / union / reference scalar).
    private readonly IReadOnlyList<ComputedValue>? _list;
    private readonly IEnumerator<ComputedValue>? _boxed;
    private RangeValueSequence.Enumerator _dense;
    private readonly bool _isDense;
    private int _index;

    private RangeValueCursor(IReadOnlyList<ComputedValue> list)
    {
        _list = list;
        _boxed = null;
        _dense = default;
        _isDense = false;
        _index = 0;
    }

    private RangeValueCursor(RangeValueSequence.Enumerator dense)
    {
        _list = null;
        _boxed = null;
        _dense = dense;
        _isDense = true;
        _index = 0;
    }

    private RangeValueCursor(IEnumerator<ComputedValue> boxed)
    {
        _list = null;
        _boxed = boxed;
        _dense = default;
        _isDense = false;
        _index = 0;
    }

    /// <summary>
    /// Opens a cursor over one argument, preferring the cheapest backing: the admitted per-epoch snapshot
    /// (zero-copy) → the dense positional stream for a closed rectangle (no allocation) → a single boxed
    /// iterator for open ranges/unions/scalars.
    /// </summary>
    public static RangeValueCursor Open(Expression argument, EvaluationContext context)
    {
        if (
            argument is Reference reference
            && context.Workbook.TryGetRangeSnapshot(reference, context) is { } snapshot
        )
        {
            return new RangeValueCursor(snapshot.Values);
        }

        switch (argument)
        {
            case RangeReference rectangle:
                return new RangeValueCursor(
                    rectangle.ExpandComputedValues(context).GetEnumerator()
                );

            case OpenRangeReference open:
                return new RangeValueCursor(open.ExpandComputedValues(context).GetEnumerator());

            case UnionReference union:
                return new RangeValueCursor(union.ExpandComputedValues(context).GetEnumerator());

            default:
                var computed = argument.Evaluate(context);
                return computed.Kind == ComputedValueKind.Reference
                    ? new RangeValueCursor(computed.EnumerateValues(context).GetEnumerator())
                    : new RangeValueCursor(Single(computed));
        }
    }

    private static IEnumerator<ComputedValue> Single(ComputedValue value)
    {
        yield return value;
    }

    /// <summary>
    /// Opens a cursor over a reference VALUE already produced by evaluating a non-reference-typed argument
    /// (e.g. a function call like OFFSET/INDEX/CHOOSE that yields a range) — the counterpart of the
    /// <c>default</c> branch of <see cref="Open"/>, but takes the already-computed value instead of
    /// re-evaluating <paramref name="computed"/>'s source expression a second time. Always the boxed-iterator
    /// backing (the same fallback a reference VALUE gets in <see cref="Open"/>): the underlying reference is
    /// walked via <see cref="ComputedValue.EnumerateValues(EvaluationContext)"/>.
    /// </summary>
    public static RangeValueCursor OpenFromReferenceValue(
        in ComputedValue computed,
        EvaluationContext context
    ) => new(computed.EnumerateValues(context).GetEnumerator());

    /// <summary>The next value in position order (column-major, matching the materialized expansion exactly),
    /// or <c>false</c> once the range is exhausted.</summary>
    public bool MoveNext(out ComputedValue value)
    {
        if (_isDense)
        {
            if (_dense.MoveNext())
            {
                value = _dense.Current;
                return true;
            }

            value = default;
            return false;
        }

        if (_list is { } list)
        {
            if (_index < list.Count)
            {
                value = list[_index++];
                return true;
            }

            value = default;
            return false;
        }

        if (_boxed!.MoveNext())
        {
            value = _boxed.Current;
            return true;
        }

        value = default;
        return false;
    }
}
