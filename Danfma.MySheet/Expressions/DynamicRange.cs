using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A ':' range whose endpoints are reference-returning EXPRESSIONS (e.g. INDEX(...):$D1), resolved to a
/// concrete <see cref="RangeReference"/> at evaluation time. The static endpoint forms (A1:B2, $D:$D) are
/// built directly by the parser and never reach this node.
/// </summary>
[MemoryPackable]
public sealed partial record DynamicRange(Expression Start, Expression End) : Reference
{
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        if (!Start.TryResolveReference(context, out var start)
            || !End.TryResolveReference(context, out var end)
            || !TryBox(start!, out var startBox)
            || !TryBox(end!, out var endBox))
        {
            return false;
        }

        var sheet = start is CellReference sc ? sc.SheetName
            : start is RangeReference sr ? sr.SheetName
            : end is CellReference ec ? ec.SheetName
            : ((RangeReference)end!).SheetName;

        var minColumn = Math.Min(startBox.MinColumn, endBox.MinColumn);
        var minRow = Math.Min(startBox.MinRow, endBox.MinRow);
        var maxColumn = Math.Max(startBox.MaxColumn, endBox.MaxColumn);
        var maxRow = Math.Max(startBox.MaxRow, endBox.MaxRow);

        reference = new RangeReference(
            new CellAddress(minColumn, minRow).ToId(),
            new CellAddress(maxColumn, maxRow).ToId(),
            sheet);
        return true;
    }

    // A ':' endpoint evaluated as a scalar mirrors a bare RangeReference: #VALUE! unless a consumer expands
    // it. Delegating to the resolved range inherits that behaviour.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryResolveReference(context, out var reference)
            ? ComputedValue.Reference(reference!)
            : ComputedValue.Error(Error.Ref);

    private readonly record struct Box(int MinColumn, int MinRow, int MaxColumn, int MaxRow);

    // Bounding box of a resolved reference. Only bounded references (cell/rectangle) are supported as a
    // dynamic endpoint; an open reference as a dynamic endpoint is out of scope (returns false -> #REF!).
    private static bool TryBox(Reference reference, out Box box)
    {
        switch (reference)
        {
            case CellReference cell:
                var a = CellAddress.Parse(cell.Id);
                box = new Box(a.Column, a.Row, a.Column, a.Row);
                return true;
            case RangeReference range:
                var s = CellAddress.Parse(range.StartId);
                var e = CellAddress.Parse(range.EndId);
                box = new Box(
                    Math.Min(s.Column, e.Column), Math.Min(s.Row, e.Row),
                    Math.Max(s.Column, e.Column), Math.Max(s.Row, e.Row));
                return true;
            default:
                box = default;
                return false;
        }
    }
}
