using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record Min(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RangeAggregate.Memoize(Arguments, context, AggregateKind.Min, () => Compute(context));

    private ComputedValue Compute(EvaluationContext context)
    {
        var fold = new MinFold();

        return NumericAggregation.Fold(Arguments, context, ref fold) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(fold.HasValue ? fold.Value : 0.0);
    }

    private struct MinFold : INumericFold
    {
        public bool HasValue;
        public double Value;

        public void Accept(double value)
        {
            if (!HasValue || value < Value)
            {
                Value = value;
                HasValue = true;
            }
        }
    }
}
