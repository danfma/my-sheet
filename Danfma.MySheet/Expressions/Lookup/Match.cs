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

        var lookup = Arguments[0].Evaluate(context);

        // Serve the lookup array from the Layer-2 range cache when the argument is a big populated range:
        // the snapshot is materialized once and every derived accelerator (exact hash, sorted prefix/suffix)
        // reproduces this scan's result bit for bit. A small range (or a non-range argument) streams the
        // memoized cells positionally (no materialized vector) via RangeValueCursor.
        var snapshot = Arguments[1] is Reference reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;

        var matchType = 1.0;

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
        if (ReferencePosition.IsLookupValueError(Arguments[0], lookup, context, out var valueError))
        {
            return valueError;
        }

        if (ReferencePosition.TryUnresolvedError(Arguments[1], context, out var unresolved))
        {
            return unresolved;
        }

        if (matchType == 0)
        {
            // Exact (type 0) → O(1) via the value→first-position hash; a blank-equivalent lookup (0/""/FALSE)
            // is the one case the hash cannot answer (Excel's intransitive blank rule) → linear fallback.
            if (snapshot is not null)
            {
                switch (snapshot.TryExactPosition(lookup, out var hashPosition))
                {
                    case ExactMatchOutcome.Found:
                        return ComputedValue.Number(hashPosition);
                    case ExactMatchOutcome.NotFound:
                        return ComputedValue.Error(Error.NA);
                }
            }

            // Threads the ALREADY-probed `snapshot` through instead of letting Open re-probe it (a snapshot
            // that answered Unsupported above is reused zero-copy; a still-null snapshot stays on its first,
            // streaming read instead of a second probe eagerly admitting it — see SUMIF's identical pattern).
            var exactPosition = 0;
            var exactCursor = RangeValueCursor.Open(Arguments[1], context, snapshot);

            while (exactCursor.MoveNext(out var value))
            {
                exactPosition++;

                if (ValueCoercion.AreEqual(value, lookup))
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

        // Approximate → O(log n) via the sorted index (correct for any input order: it returns the LAST
        // position among the qualifying values, exactly like the linear scan below).
        if (snapshot is not null)
        {
            var indexed =
                matchType > 0
                    ? snapshot.ApproximateAscendingPosition(lookup)
                    : snapshot.ApproximateDescendingPosition(lookup);

            return indexed >= 1 ? ComputedValue.Number(indexed) : ComputedValue.Error(Error.NA);
        }

        // `snapshot` is guaranteed null here (a non-null snapshot always returns above), so threading it
        // through — rather than letting Open re-probe — keeps this the range's first, streaming read.
        var position = -1;
        var index = 0;
        var approxCursor = RangeValueCursor.Open(Arguments[1], context, snapshot);

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
}
