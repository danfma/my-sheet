using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record IfNa(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = Arguments[0].Evaluate(context);

        // Only #N/A is caught; other errors pass through.
        return value.TryGetError(out var error) && error == Error.NA
            ? Arguments[1].Evaluate(context)
            : value;
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
