using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RangeAggregate.Memoize(Arguments, context, AggregateKind.CountA, () => Compute(context));

    private ComputedValue Compute(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF!, not an empty range that COUNTA would tally 0.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var count = 0;

        foreach (var value in ArgumentFlattening.FlattenComputedValues(Arguments, context))
        {
            if (value.Kind != ComputedValueKind.Blank)
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
