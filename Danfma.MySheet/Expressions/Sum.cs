using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Sum(Expression[] Expressions) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new SumFold();

        return NumericAggregation.Fold(Expressions, context, ref fold) is { } error
            ? ComputedValue.From(error)
            : ComputedValue.Number(fold.Total);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();

    private struct SumFold : INumericFold
    {
        public double Total;

        public void Accept(double value) => Total += value;
    }
}
