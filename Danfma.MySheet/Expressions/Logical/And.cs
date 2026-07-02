using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record And(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var result = true;

        foreach (var argument in Arguments)
        {
            if (argument.Evaluate(context).CoerceToBool(out var value) is { } error)
            {
                return ComputedValue.Error(error);
            }

            result &= value;
        }

        return ComputedValue.Boolean(result);
    }
}
