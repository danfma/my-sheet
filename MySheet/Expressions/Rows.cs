using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Rows(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook) =>
        Arguments[0] is RangeReference range ? (double)range.RowCount : 1.0;
}
