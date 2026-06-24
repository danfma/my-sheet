using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record IfError(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var value = Arguments[0].Compute(context);

        // The fallback is only computed when the first argument is an error (short-circuit).
        return value is ErrorValue ? Arguments[1].Compute(context) : value;
    }
}
