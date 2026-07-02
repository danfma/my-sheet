using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record SumIf(Expression[] Arguments) : Function
{
    // SUMIF(range, criteria, [sum_range]) — sums sum_range (or range) where range matches the criteria.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));
        var range = ArgumentFlattening.ExpandComputedValues(Arguments[0], context);
        var sumRange =
            Arguments.Length == 3 ? ArgumentFlattening.ExpandComputedValues(Arguments[2], context) : range;

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
