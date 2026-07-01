using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record NumberValue(double Value) : ValueExpression
{
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Number(Value);

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
