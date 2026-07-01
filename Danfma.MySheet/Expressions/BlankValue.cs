using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record BlankValue : ValueExpression
{
    public static readonly BlankValue Instance = new();

    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Blank;

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
