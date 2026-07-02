using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record Not(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToBool(out var value) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Boolean(!value);
    }
}
