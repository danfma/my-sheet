using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record StringValue(string Value) : ValueExpression
{
    public override object? Compute(Workbook workbook) => Value;
}