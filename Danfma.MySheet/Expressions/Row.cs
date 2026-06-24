using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Row(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context) =>
        Arguments switch
        {
            [CellReference cell] => (double)CellAddress.Parse(cell.Id).Row,
            [RangeReference range] => (double)range.TopRow,
            // ROW() with no argument uses the cell currently being evaluated, when one is known.
            [] when context.CellId is { } id => (double)CellAddress.Parse(id).Row,
            _ => ErrorValue.NotValue,
        };
}
