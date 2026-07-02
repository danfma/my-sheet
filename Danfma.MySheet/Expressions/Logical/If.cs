using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record If(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToBool(out var condition) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Only the taken branch is computed (short-circuit), matching Excel.
        if (condition)
        {
            return Arguments[1].Evaluate(context);
        }

        return Arguments.Length == 3 ? Arguments[2].Evaluate(context) : ComputedValue.Boolean(false);
    }
}
