using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Abs(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var number) is { } error)
        {
            return error;
        }

        return Math.Abs(number);
    }
}
