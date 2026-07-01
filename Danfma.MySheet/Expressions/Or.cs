using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Or(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var result = false;

        foreach (var argument in Arguments)
        {
            if (argument.Evaluate(context).CoerceToBool(out var value) is { } error)
            {
                return ComputedValue.Error(error);
            }

            result |= value;
        }

        return ComputedValue.Boolean(result);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
