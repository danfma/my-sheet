using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Sum(Expression[] Expressions) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var fold = new SumFold();
        var error = NumericAggregation.Fold(Expressions, context, ref fold);

        return error ?? (object)fold.Total;
    }

    private struct SumFold : INumericFold
    {
        public double Total;

        public void Accept(double value) => Total += value;
    }
}
