using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record IsNumber(Expression[] Arguments) : Function
{
    // An error value is "not a number" rather than propagated, matching Excel's IS functions.
    public override object? Compute(EvaluationContext context) => Arguments[0].Compute(context) is double;
}
