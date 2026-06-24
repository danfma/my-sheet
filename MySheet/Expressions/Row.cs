using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Row(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook) => Arguments switch
    {
        [CellReference cell] => (double)CellAddress.Parse(cell.Id).Row,
        [RangeReference range] => (double)range.TopRow,
        // ROW() with no argument needs the formula's own cell, which the evaluator does not track yet.
        _ => ErrorValue.NotValue,
    };
}
