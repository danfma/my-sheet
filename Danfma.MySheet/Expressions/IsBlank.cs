using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record IsBlank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind == ComputedValueKind.Blank);

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
