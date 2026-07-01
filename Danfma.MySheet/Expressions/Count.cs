using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Count(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // COUNT only tallies numeric values and, unlike SUM, never propagates errors.
        var fold = new CountFold();
        NumericAggregation.Fold(Arguments, context, ref fold);

        return ComputedValue.Number(fold.Count);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();

    private struct CountFold : INumericFold
    {
        public int Count;

        public void Accept(double value) => Count++;
    }
}
