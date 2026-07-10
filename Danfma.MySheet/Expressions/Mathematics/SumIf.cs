using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record SumIf(Expression[] Arguments) : Function
{
    // SUMIF(range, criteria, [sum_range]) — sums sum_range (or range) where range matches the criteria.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A range (or sum_range) over a missing sheet is a structural #REF!, not an empty range summing 0.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));
        var range = ArgumentFlattening.ExpandCached(Arguments[0], context, out var snapshot);

        // Single-range (no separate sum_range) numeric `=k` criterion → O(1) via the numeric-equality map,
        // whose per-key Sum is exactly Σ of the matching cells' own numeric values. Any other shape scans
        // the cached snapshot(s) linearly (no cell re-read).
        if (
            Arguments.Length < 3
            && snapshot is not null
            && criteria.TryGetNumericEquality(out var key)
        )
        {
            return ComputedValue.Number(snapshot.NumericEquality(key).Sum);
        }

        var sumRange =
            Arguments.Length == 3
                ? ArgumentFlattening.ExpandCached(Arguments[2], context, out _)
                : range;

        var total = 0.0;
        var length = Math.Min(range.Count, sumRange.Count);

        for (var i = 0; i < length; i++)
        {
            if (criteria.Matches(range[i]) && sumRange[i].TryGetNumber(out var number))
            {
                total += number;
            }
        }

        return ComputedValue.Number(total);
    }
}
