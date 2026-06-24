using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Value(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        // Reuse numeric coercion: numeric text parses, non-numeric text → #VALUE!.
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var number) is { } error)
        {
            return error;
        }

        return number;
    }
}
