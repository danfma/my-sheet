using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Len(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(context), out var text) is { } error)
        {
            return error;
        }

        return (double)text.Length;
    }
}
