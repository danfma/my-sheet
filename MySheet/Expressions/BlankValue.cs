using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record BlankValue : ValueExpression
{
    public static readonly BlankValue Instance = new();

    public override object? Compute(EvaluationContext context) => null;
}