using MemoryPack;

namespace Danfma.MySheet.Expressions.Information;

[MemoryPackable]
public sealed partial record IsNumber(Expression[] Arguments) : Function
{
    // An error value is "not a number" rather than propagated, matching Excel's IS functions.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind == ComputedValueKind.Number);

}
