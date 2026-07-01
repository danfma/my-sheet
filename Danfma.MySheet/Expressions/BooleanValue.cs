using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record BooleanValue(bool Value) : ValueExpression
{
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Boolean(Value);

}
