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
        // reproduces this scan's result bit for bit. A small range (or a non-range argument) keeps the
        // original linear enumeration.
        var snapshot = Arguments[1] is Reference reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;
        IReadOnlyList<ComputedValue> array =
            snapshot?.Values
            ?? (IReadOnlyList<ComputedValue>)ArgumentFlattening.ExpandComputedValues(Arguments[1], context);

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

            for (var i = 0; i < array.Count; i++)
            {
                if (ValueCoercion.AreEqual(array[i], lookup))
                {
                    return ComputedValue.Number(i + 1);
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
            var indexed = matchType > 0
                ? snapshot.ApproximateAscendingPosition(lookup)
                : snapshot.ApproximateDescendingPosition(lookup);

            return indexed >= 1 ? ComputedValue.Number(indexed) : ComputedValue.Error(Error.NA);
        }

        var position = -1;

        for (var i = 0; i < array.Count; i++)
        {
            var value = array[i];
            if (value.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
            {
                continue;
            }

            var comparison = ValueCoercion.Compare(value, lookup);

            if (matchType > 0 && comparison <= 0)
            {
                position = i + 1;
            }
            else if (matchType < 0 && comparison >= 0)
            {
                position = i + 1;
            }
        }

        return position >= 1 ? ComputedValue.Number(position) : ComputedValue.Error(Error.NA);
    }
}
