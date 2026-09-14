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
    // (closed rectangle) or a boxed iterator (open range / union / reference scalar). When set, the
    // argument's OWN value was an error (sweep item 34a) and no backing is live.
    private readonly IReadOnlyList<ComputedValue>? _list;
    private readonly IEnumerator<ComputedValue>? _boxed;
    private readonly Error? _slotError;
    private RangeValueSequence.Enumerator _dense;
    private readonly bool _isDense;
    private int _index;

    /// <summary>
    /// The argument's own error, sweep item 34(a) — the mirror of
    /// <see cref="PositionalRange.SlotError"/>: an error-valued argument in the criteria family's range
    /// slot (COUNTIF is this cursor's member) propagates instead of streaming as the one element the
    /// criteria discards. Set only by the fallback of <see cref="Open(Expression, EvaluationContext,
    /// RangeSnapshot?)"/>; the consumer checks it right after opening, before any scan.
    /// </summary>
    public readonly Error? SlotError => _slotError;

    private RangeValueCursor(IReadOnlyList<ComputedValue> list)
    {
        _list = list;
        _boxed = null;
        _slotError = null;
        _dense = default;
        _isDense = false;
        _index = 0;
    }

    private RangeValueCursor(RangeValueSequence.Enumerator dense)
    {
        _list = null;
        _boxed = null;
        _slotError = null;
        _dense = dense;
        _isDense = true;
        _index = 0;
    }

    private RangeValueCursor(IEnumerator<ComputedValue> boxed)
    {
        _list = null;
        _boxed = boxed;
        _slotError = null;
        _dense = default;
        _isDense = false;
        _index = 0;
    }

    private RangeValueCursor(Error slotError)
    {
        _list = null;
        _boxed = null;
        _slotError = slotError;
        _dense = default;
        _isDense = false;
        _index = 0;
    }

    /// <summary>
    /// Opens a cursor over one argument, probing the range cache itself, preferring the cheapest backing: the
    /// admitted per-epoch snapshot (zero-copy) → the dense positional stream for a closed rectangle (no
    /// allocation) → a single boxed iterator for open ranges/unions/scalars.
    /// </summary>
    public static RangeValueCursor Open(Expression argument, EvaluationContext context) =>
        Open(
            argument,
            context,
            argument is Reference reference
                ? context.Workbook.TryGetRangeSnapshot(reference, context)
                : null
        );

    /// <summary>
    /// Same backing preference as <see cref="Open(Expression, EvaluationContext)"/>, but takes an
    /// ALREADY-resolved snapshot probe instead of running its own. <see cref="Workbook.TryGetRangeSnapshot"/>
    /// is stateful (second-use admission): calling it a SECOND time for the same range within one function
    /// evaluation — e.g. once for an exact/approximate hash-path check, then again here — would itself count
    /// as the admitting "second read" and eagerly build the snapshot on what should still be the range's
    /// first, streaming read. A caller that already probed the snapshot (COUNTIF/MATCH/XLOOKUP's fast-path
    /// check) MUST route that same result through here rather than calling the no-snapshot overload again.
    /// Mirrors <see cref="PositionalRange.Open(Expression, EvaluationContext, RangeSnapshot?)"/>.
    /// </summary>
    public static RangeValueCursor Open(
        Expression argument,
        EvaluationContext context,
        RangeSnapshot? snapshot
    )
    {
        if (snapshot is not null)
        {
            return new RangeValueCursor(snapshot.Values);
        }

        // Phase 2 audit (shared-formula delta production): mirrors PositionalRange.Open's identical fix — a
        // COUNTIF/MATCH/XLOOKUP range argument written INSIDE a shared-formula master is an anchored node;
        // resolving it up front lets the switch below (and its RangeReference fast path) treat it exactly
        // like an ordinary range, with no logic duplicated.
        if (argument is AnchoredRangeReference anchoredRange)
        {
            argument = anchoredRange.ToRangeReference(context);
        }

        // Phase 5 item 16, PERF only: resolves a TableReference to its concrete rectangle UP FRONT, same
        // shape as the AnchoredRangeReference resolution above, so the switch's RangeReference fast path
        // (the allocation-free struct enumerator) serves a resolved table too. An unresolvable table is
        // left as-is: it falls through to `default:` below, whose sweep 34(a) arm carries the node's own
        // #NAME?/#REF! on <see cref="SlotError"/> for COUNTIF — since TryStream/IsBareReferenceNode never
        // enter this method at all. An EMPTY area (sweep item 33) is left as-is too: its value is an
        // EmptyRangeReference, which the `default:` arm streams as nothing (MATCH/XLOOKUP #N/A, COUNTIF 0).
        if (
            argument is TableReference table
            && table.TryResolveRectangle(context.Workbook, out var tableRange)
        )
        {
            argument = tableRange;
        }

        var referenceReturningNode = NamedReferences.TryResolveReferenceReturningNode(
            argument,
            context,
            out var selectedReference,
            out var unresolvedValue
        );
        if (referenceReturningNode == NamedReferences.ReferenceReturningNodeResolution.Resolved)
        {
            argument = selectedReference;
        }
        else if (
            referenceReturningNode == NamedReferences.ReferenceReturningNodeResolution.Unresolved
        )
        {
            return new RangeValueCursor(
                unresolvedValue.TryGetError(out var error) ? error : Error.Ref
            );
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
                // Sweep item 34(a), mirroring PositionalRange.Open's fallback arm: an argument whose OWN
                // value is an error (PositionalRange.IsOwnSlotError carries the guard — a cell reference
                // to an error cell is content, not the argument's error) is carried on
                // <see cref="SlotError"/> for the criteria family's COUNTIF to surface, instead of
                // streaming as the one element the criteria discards (a silent 0 where the oracle answers
                // the node's #NAME?/#DIV/0!/#REF! — measured 2026-09-11, both entry modes). The scan
                // consumers (MATCH/XLOOKUP) do not read it: they see the same single-element stream as
                // before, and the unresolved-NAME column of the oracle keeps their own codes (see Match.cs).
                var computed = argument.Evaluate(context);

                if (
                    computed.TryGetError(out var slotError)
                    && PositionalRange.IsOwnSlotError(argument, slotError, context)
                )
                {
                    return new RangeValueCursor(slotError);
                }

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
    /// or <c>false</c> once the range is exhausted. The error-carrying cursor (sweep item 34(a)) yields the
    /// error as its ONE element, so a consumer that does not surface <see cref="SlotError"/> sees exactly the
    /// stream this cursor produced before the arm existed.</summary>
    public bool MoveNext(out ComputedValue value)
    {
        if (_slotError is { } slotError)
        {
            if (_index++ == 0)
            {
                value = ComputedValue.Error(slotError);
                return true;
            }

            value = default;
            return false;
        }

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
