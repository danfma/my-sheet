using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Upper(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(context), out var text) is { } error)
        {
            return error;
        }

        return text.ToUpperInvariant();
    }
}
