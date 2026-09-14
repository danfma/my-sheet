using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record IfNa(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = ScalarReferenceValue.Evaluate(Arguments[0], context);

        // Only #N/A is caught; other errors pass through.
        return value.TryGetError(out var error) && error == Error.NA
            ? Arguments[1].Evaluate(context)
            : value;
    }
}
