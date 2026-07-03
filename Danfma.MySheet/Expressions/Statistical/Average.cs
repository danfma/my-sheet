using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record Average(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RangeAggregate.Memoize(Arguments, context, AggregateKind.Average, () => Compute(context));

    private ComputedValue Compute(EvaluationContext context)
    {
        var fold = new AverageFold();

        if (NumericAggregation.Fold(Arguments, context, ref fold) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return fold.Count == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(fold.Total / fold.Count);
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
