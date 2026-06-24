using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record And(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var result = true;

        foreach (var argument in Arguments)
        {
            if (ValueCoercion.TryToBool(argument.Compute(context), out var value) is { } error)
            {
                return error;
            }

            result &= value;
        }

        return result;
    }
}
