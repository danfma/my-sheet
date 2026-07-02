using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record Max(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new MaxFold();

        return NumericAggregation.Fold(Arguments, context, ref fold) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(fold.HasValue ? fold.Value : 0.0);
    }

    private struct MaxFold : INumericFold
    {
        public bool HasValue;
        public double Value;

        public void Accept(double value)
        {
            if (!HasValue || value > Value)
            {
                Value = value;
                HasValue = true;
            }
        }
    }
}
