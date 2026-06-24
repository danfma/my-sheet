using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record BooleanValue(bool Value) : ValueExpression
{
    public override object? Compute(Workbook workbook) => Value;
}