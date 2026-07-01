using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountIf(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
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
