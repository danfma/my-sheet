using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountIf(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A range over a missing sheet is a structural #REF!, not an empty range that matches nothing.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));

        var snapshot = Arguments[0] is Reference reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;

        // Admitted range (big, populated) → serve the shared snapshot: a numeric `=k` criterion is O(1) via
        // the numeric-equality map; anything else (text/wildcard/comparison) scans the snapshot's array
        // linearly (zero-copy, no cell re-read).
        if (snapshot is not null)
        {
            if (criteria.TryGetNumericEquality(out var key))
            {
                return ComputedValue.Number(snapshot.NumericEquality(key).Count);
            }

            var matches = 0;

            foreach (var value in snapshot.Values)
            {
                if (criteria.Matches(value))
                {
                    matches++;
                }
            }

            return ComputedValue.Number(matches);
        }

        // Non-admitted range → stream the memoized cells positionally (dense struct cursor for a closed
        // rectangle, one small boxed iterator for an open range/union) instead of materializing the whole
        // vector just to count it. Threads the ALREADY-probed `snapshot` (null here, since a non-null
        // snapshot returns above) through instead of letting Open re-probe it: TryGetRangeSnapshot is the
        // second-use ADMISSION check itself, so a second call here would eagerly build the snapshot on what
        // must stay this range's first, streaming read — see SUMIF's identical pattern.
        var count = 0;
        var cursor = RangeValueCursor.Open(Arguments[0], context, snapshot);

        while (cursor.MoveNext(out var value))
        {
            if (criteria.Matches(value))
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
