using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Row(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Arguments switch
        {
            [CellReference cell] => ComputedValue.Number(CellAddress.Parse(cell.Id).Row),
            [RangeReference range] => ComputedValue.Number(range.TopRow),
            // ROW() with no argument uses the cell currently being evaluated, when one is known.
            [] when context.CellId is { } id => ComputedValue.Number(CellAddress.Parse(id).Row),
            _ => ComputedValue.Error(Error.Value),
        };

}
