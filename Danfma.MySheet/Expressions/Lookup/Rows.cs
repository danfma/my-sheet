using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Rows(Expression[] Arguments) : Function
{
    // A defined name that stands for a range counts its rows; a whole-column/row reference uses the
    // populated extent on its open row axis (structural on a bounded one); anything else (a single cell
    // or a scalar) is 1. boundOpenRanges:false keeps the open reference so the extent rule applies.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Number(
            NamedReferences.TryResolveReference(Arguments[0], context, out var reference, boundOpenRanges: false)
                ? reference switch
                {
                    RangeReference range => range.RowCount,
                    OpenRangeReference open => open.RowExtent(context),
                    _ => 1.0,
                }
                : 1.0
        );
}
