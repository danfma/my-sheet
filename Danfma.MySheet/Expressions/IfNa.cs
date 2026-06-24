using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record IfNa(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var value = Arguments[0].Compute(context);

        // Only #N/A is caught; other errors pass through.
        return value is ErrorValue { ErrorCode: "#N/A" } ? Arguments[1].Compute(context) : value;
    }
}
