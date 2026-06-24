using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Average(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var fold = new AverageFold();
        var error = NumericAggregation.Fold(Arguments, context, ref fold);

        if (error is not null)
        {
            return error;
        }

        return fold.Count == 0 ? ErrorValue.DivByZero : fold.Total / fold.Count;
    }

    private struct AverageFold : INumericFold
    {
        public double Total;
        public int Count;

        public void Accept(double value)
        {
            Total += value;
            Count++;
        }
    }
}
