using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record If(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToBool(Arguments[0].Compute(context), out var condition) is { } error)
        {
            return error;
        }

        // Only the taken branch is computed (short-circuit), matching Excel.
        if (condition)
        {
            return Arguments[1].Compute(context);
        }

        return Arguments.Length == 3 ? Arguments[2].Compute(context) : (object)false;
    }
}
