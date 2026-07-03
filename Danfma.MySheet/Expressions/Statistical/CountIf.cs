using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountIf(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A range over a missing sheet is a structural #REF!, not an empty range that matches nothing.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));
        var count = 0;

        foreach (var value in ArgumentFlattening.ExpandComputedValues(Arguments[0], context))
        {
            if (criteria.Matches(value))
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
