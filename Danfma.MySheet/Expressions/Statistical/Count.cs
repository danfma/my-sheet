using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record Count(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RangeAggregate.Memoize(Arguments, context, AggregateKind.Count, () => Compute(context));

    private ComputedValue Compute(EvaluationContext context)
    {
        // COUNT ignores cell VALUE errors, but a reference to a missing sheet is a STRUCTURAL #REF! it must
        // still surface — so it is guarded explicitly (COUNT discards Fold's error channel).
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // COUNT only tallies numeric values and, unlike SUM, never propagates errors.
        var fold = new CountFold();
        NumericAggregation.Fold(Arguments, context, ref fold);

        return ComputedValue.Number(fold.Count);
    }

    private struct CountFold : INumericFold
    {
        public int Count;

        public void Accept(double value) => Count++;
    }
}
