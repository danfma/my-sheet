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
        // keys sort lexicographically, exactly like the <= operator — not only numeric keys.
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
