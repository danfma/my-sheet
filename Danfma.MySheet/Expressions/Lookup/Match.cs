using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Match(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet lookup array is a structural #REF! — distinct from an empty array over an existing
        // sheet, which stays #N/A. Guard before enumerating so the missing sheet is not swallowed as empty.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var lookup = LookupMatching.EvaluateKey(
            Arguments[0],
            context,
            out var absentLookup,
            out var lookupReference
        );

        var matchType = 1.0;

        Reference? arrayReference = null;
        _ = NamedReferences.TryResolveReference(
            Arguments[1],
            context,
            out arrayReference,
            boundOpenRanges: false
        );
        var open = arrayReference as OpenRangeReference;
        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out matchType) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        // Sweep item 34(b), both slots in one place (after the match-type parse, whose own coercion error
        // keeps its precedence): the lookup VALUE's error leads the scan on both match-type paths — its own
        // error, or a single cell's that holds one (ReferencePosition.IsLookupValueError: MATCH(A1,A1) over
        // A1 = =1/0 is #DIV/0! for every match type on Aspose.Cells 26.6.0 and 26.7.0) — and when the lookup
        // ARRAY's node does not resolve, the node's OWN error leads instead of the not-found #N/A (#NAME? for
        // an unknown name, the node's #REF! for an unresolvable structured reference; the oracle answers the
        // error on every one of these shapes, both entry modes).
        if (
            ReferencePosition.IsLookupValueError(
                Arguments[0],
                lookup,
                lookupReference,
                context,
                out var valueError
            )
        )
        {
            return valueError;
        }

        if (
            arrayReference is null
            && (
                Arguments[1] is NameReference or TableReference
                || ArrayEvaluation.IsArrayEligible(Arguments[1], context)
                || Arguments[1] is ErrorValue
            )
            && ReferencePosition.TryUnresolvedError(Arguments[1], context, out var unresolved)
        )
        {
            return unresolved;
        }

        // Sweep item 37 follow-up, ruling (a): an OPEN lookup array ($4:$4, $5:$1000, A:A, …) returns an
        // ABSOLUTE position — column A / row 1 is position 1, never "the Nth POPULATED cell" the ordinary
        // value-only RangeValueCursor counts (it has no per-cell coordinate to translate, only values).
        // boundOpenRanges: false keeps it open here instead of collapsing to the populated bounding box's
        // own corner. Bypasses the snapshot/RangeValueCursor machinery entirely for this shape — it needs
        // each populated cell's OWN coordinate, which neither carries — but still only ever visits
        // POPULATED cells via the structural index (OpenRangeReference.PopulatedCells), never the whole
        // grid, so the "never materialise the full row" contract holds.
        // Serve the lookup array from the Layer-2 range cache when the argument is a big populated range:
        // the snapshot is materialized once and every derived accelerator (exact hash, sorted prefix/suffix)
        // reproduces this scan's result bit for bit. A small range (or a non-range argument) streams the
        // memoized cells positionally (no materialized vector) via RangeValueCursor.
        var snapshot = arrayReference is { } reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;

        if (open is not null && snapshot is null)
        {
            return MatchOverOpenRange(open, lookup, matchType, context);
        }

        if (matchType == 0)
        {
            if (LookupMatching.UsesWildcards(lookup))
            {
                var wildcardCursor = RangeValueCursor.Open(
                    arrayReference ?? Arguments[1],
                    context,
                    snapshot
                );
                var wildcardValues = new List<ComputedValue>();
                while (wildcardCursor.MoveNext(out var value))
                {
                    wildcardValues.Add(value);
                }

                var wildcardPosition = LookupMatching.FindMatch(
                    lookup,
                    wildcardValues,
                    wildcardValues.Count,
                    matchMode: 2,
                    reverse: false
                );
                if (wildcardPosition < 0)
                {
                    return ComputedValue.Error(Error.NA);
                }

                var oneBasedPosition = wildcardPosition + 1;
                return ComputedValue.Number(
                    snapshot?.SourcePosition(oneBasedPosition) ?? oneBasedPosition
                );
            }

            // Exact (type 0) → O(1) via the value→first-position hash; a blank-equivalent lookup (0/""/FALSE)
            // is the one case the hash cannot answer (Excel's intransitive blank rule) → linear fallback.
            if (snapshot is not null)
            {
                switch (snapshot.TryExactPosition(lookup, out var hashPosition))
                {
                    case ExactMatchOutcome.Found:
                        return ComputedValue.Number(snapshot.SourcePosition(hashPosition));
                    case ExactMatchOutcome.NotFound:
                        return ComputedValue.Error(Error.NA);
                }
            }

            // Threads the ALREADY-probed `snapshot` through instead of letting Open re-probe it (a snapshot
            // that answered Unsupported above is reused zero-copy; a still-null snapshot stays on its first,
            // streaming read instead of a second probe eagerly admitting it — see SUMIF's identical pattern).
            var exactPosition = 0;
            var exactCursor = RangeValueCursor.Open(
                arrayReference ?? Arguments[1],
                context,
                snapshot
            );
            var exactMatcher = new LookupMatching.ExactMatcher(lookup);

            while (exactCursor.MoveNext(out var value))
            {
                exactPosition++;

                if (exactMatcher.Matches(value))
                {
                    return ComputedValue.Number(exactPosition);
                }
            }

            return ComputedValue.Error(Error.NA);
        }

        // Approximate: matchType > 0 assumes ascending (largest value <= lookup); < 0 assumes
        // descending (smallest value >= lookup). Cross-type ordering (ValueCoercion.Compare) lets text
        // keys sort lexicographically, exactly the <= operator — not only numeric keys.

        // Final-review fix wave, finding I1: main's UNCONDITIONAL rule on this path — ANY lookup-value
        // error leads, a range-collapse #VALUE! included (`MATCH(B1:B3,B1:B3)` is #VALUE!, both main's and
        // the oracle PLAIN's answer, Aspose.Cells 26.7.0) — restored here, verbatim. T1 (`a542f56`) deleted
        // it and replaced the one check above (IsLookupValueError) as the ONLY guard on both paths;
        // IsLookupValueError answers false for a RangeReference (its own #VALUE! is "content", the rule
        // PositionalRange.IsOwnSlotError was written for), so a range-collapse lookup value fell through to
        // the scan below and answered a POSITION instead of propagating. The EXACT path above deliberately
        // keeps ONLY IsLookupValueError's narrower guard: a range's #VALUE! there is the pinned collapse
        // artifact (`SUM(MATCH(A5:A7,A5:A7,0))` stays #N/A, ElementwiseLiftingTests' known divergence) —
        // this unconditional check must not reach it. XMATCH and XLOOKUP do not need the same restoration:
        // main never propagated a lookup-value error for either (measured against `a02ed5d`'s
        // LookupFunctions.cs/XLookup.cs — no such check existed at all, on any match mode), so their
        // current IsLookupValueError-only guard is a net addition over main, not a regression.
        if (lookup.Kind == ComputedValueKind.Error)
        {
            return lookup;
        }

        if (lookup.TryGetText(out var lookupText) && lookupText.Length == 0)
        {
            var textPosition = 0;
            var textMatch = -1;
            var textCursor = RangeValueCursor.Open(
                arrayReference ?? Arguments[1],
                context,
                snapshot
            );
            while (textCursor.MoveNext(out var value))
            {
                textPosition++;
                if (LookupMatching.IsExactText(value, lookupText))
                {
                    if (matchType < 0)
                    {
                        return ComputedValue.Number(textPosition);
                    }

                    textMatch = textPosition;
                }
            }

            return textMatch >= 0 ? ComputedValue.Number(textMatch) : ComputedValue.Error(Error.NA);
        }

        if (absentLookup)
        {
            var absentPosition = 0;
            var absentCursor = RangeValueCursor.Open(
                arrayReference ?? Arguments[1],
                context,
                snapshot
            );
            while (absentCursor.MoveNext(out var value))
            {
                absentPosition++;
                if (ValueCoercion.AreEqual(value, ComputedValue.Number(0)))
                {
                    return ComputedValue.Number(absentPosition);
                }
            }

            return ComputedValue.Error(Error.NA);
        }

        // Approximate → O(log n) via the sorted index (correct for any input order: it returns the LAST
        // position among the qualifying values, exactly like the linear scan below).
        if (snapshot is not null)
        {
            var indexed =
                matchType > 0
                    ? snapshot.ApproximateAscendingPosition(lookup)
                    : snapshot.ApproximateDescendingPosition(lookup);

            return indexed >= 1
                ? ComputedValue.Number(snapshot.SourcePosition(indexed))
                : ComputedValue.Error(Error.NA);
        }

        // `snapshot` is guaranteed null here (a non-null snapshot always returns above), so threading it
        // through — rather than letting Open re-probe — keeps this the range's first, streaming read.
        var position = -1;
        var index = 0;
        var approxCursor = RangeValueCursor.Open(arrayReference ?? Arguments[1], context, snapshot);

        while (approxCursor.MoveNext(out var value))
        {
            index++;

            if (value.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
            {
                continue;
            }

            var comparison = ValueCoercion.Compare(value, lookup);

            if (matchType > 0 && comparison <= 0)
            {
                position = index;
            }
            else if (matchType < 0 && comparison >= 0)
            {
                position = index;
            }
        }

        return position >= 1 ? ComputedValue.Number(position) : ComputedValue.Error(Error.NA);
    }

    // Sweep item 37 follow-up, ruling (a): MATCH over a genuinely OPEN lookup array. Walks
    // OpenRangeReference.PopulatedCells directly (column/row pairs, not just values) so each visited
    // cell's OWN absolute position — IsSingleRow selects the COLUMN axis (a whole-ROW array like $4:$4),
    // otherwise the ROW axis (a whole-COLUMN array) — replaces the ordinal counter the closed-range scan
    // above uses. Exact (type 0) returns the FIRST match in scan order; approximate keeps the LAST
    // qualifying position in scan order, mirroring the closed-range linear scan's own rule (Excel's
    // approximate match is "last element satisfying the condition in array order", not nearest-value).
    private static ComputedValue MatchOverOpenRange(
        OpenRangeReference open,
        ComputedValue lookup,
        double matchType,
        EvaluationContext context
    )
    {
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(open.SheetName);

        int PositionOf(int column, int row) =>
            open.IsSingleRow ? open.ColumnPosition(column) : open.RowPosition(row);

        if (matchType == 0)
        {
            foreach (var (column, row) in open.PopulatedCells(context))
            {
                var value = workbook.GetCellValueDense(handle, open.SheetName, column, row);

                if (ValueCoercion.AreEqual(value, lookup))
                {
                    return ComputedValue.Number(PositionOf(column, row));
                }
            }

            return ComputedValue.Error(Error.NA);
        }

        var position = -1;

        foreach (var (column, row) in open.PopulatedCells(context))
        {
            var value = workbook.GetCellValueDense(handle, open.SheetName, column, row);

            if (value.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
            {
                continue;
            }

            var comparison = ValueCoercion.Compare(value, lookup);

            if (matchType > 0 && comparison <= 0)
            {
                position = PositionOf(column, row);
            }
            else if (matchType < 0 && comparison >= 0)
            {
                position = PositionOf(column, row);
            }
        }

        return position >= 1 ? ComputedValue.Number(position) : ComputedValue.Error(Error.NA);
    }
}
