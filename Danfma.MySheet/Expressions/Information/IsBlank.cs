using MemoryPack;

namespace Danfma.MySheet.Expressions.Information;

[MemoryPackable]
public sealed partial record IsBlank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Boolean(Arguments[0].Evaluate(context).Kind == ComputedValueKind.Blank);
}
