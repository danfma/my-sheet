using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record NumberValue(double Value) : ValueExpression
{
    public override object? Compute(EvaluationContext context) => Value;
}
