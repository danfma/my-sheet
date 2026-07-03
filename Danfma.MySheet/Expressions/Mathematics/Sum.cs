using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record Sum(Expression[] Expressions) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RangeAggregate.Memoize(Expressions, context, AggregateKind.Sum, () => Compute(context));

    private ComputedValue Compute(EvaluationContext context)
    {
        var fold = new SumFold();

        return NumericAggregation.Fold(Expressions, context, ref fold) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(fold.Total);
    }

    private struct SumFold : INumericFold
    {
        public double Total;

        public void Accept(double value) => Total += value;
    }
}
