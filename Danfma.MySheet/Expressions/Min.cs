using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Min(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new MinFold();

        return NumericAggregation.Fold(Arguments, context, ref fold) is { } error
            ? ComputedValue.From(error)
            : ComputedValue.Number(fold.HasValue ? fold.Value : 0.0);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();

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
