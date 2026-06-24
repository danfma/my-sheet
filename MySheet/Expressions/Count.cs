using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Count(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        // COUNT only tallies numeric values and, unlike SUM, never propagates errors.
        var fold = new CountFold();
        NumericAggregation.Fold(Arguments, context, ref fold);

        return (double)fold.Count;
    }

    private struct CountFold : INumericFold
    {
        public int Count;

        public void Accept(double value) => Count++;
    }
}
