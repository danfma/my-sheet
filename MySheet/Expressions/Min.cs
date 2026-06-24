using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Min(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var fold = new MinFold();
        var error = NumericAggregation.Fold(Arguments, context, ref fold);

        return error ?? (object)(fold.HasValue ? fold.Value : 0.0);
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
