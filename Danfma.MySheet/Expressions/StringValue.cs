using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record StringValue(string Value) : ValueExpression
{
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Text(Value);

}
