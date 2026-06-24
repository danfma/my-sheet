using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record StringValue(string Value) : ValueExpression
{
    public override object? Compute(EvaluationContext context) => Value;
}
