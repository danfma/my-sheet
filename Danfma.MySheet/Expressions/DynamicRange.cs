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

        if (
            !Start.TryResolveReference(context, out var start)
            || !End.TryResolveReference(context, out var end)
            || !TryBox(start!, out var startBox)
            || !TryBox(end!, out var endBox)
        )
        {
            return false;
        }

        var startSheet = SheetOf(start!);
        var endSheet = SheetOf(end!);

        // A cross-sheet range (Sheet1!A1:Sheet2!B2) is #REF! in Excel, not a silent resolve against the
        // start endpoint's sheet. Both endpoints must live on the same sheet.
        if (!string.Equals(startSheet, endSheet, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var minColumn = Math.Min(startBox.MinColumn, endBox.MinColumn);
        var minRow = Math.Min(startBox.MinRow, endBox.MinRow);
        var maxColumn = Math.Max(startBox.MaxColumn, endBox.MaxColumn);
        var maxRow = Math.Max(startBox.MaxRow, endBox.MaxRow);

        // Sweep item 33: two zero-row endpoints (Tabela1[Valor]:Tabela1[Qtd] over a header-only table) span a
        // box whose bottom is above its top — a ZERO-row rectangle (oracle: ROWS 0, COLUMNS 2, both modes),
        // which only an EmptyRangeReference can carry; a RangeReference would normalize it onto the header.
        reference =
            maxRow < minRow
                ? new EmptyRangeReference(startSheet, minRow, minColumn, maxColumn)
                : new RangeReference(
                    new CellAddress(minColumn, minRow).ToId(),
                    new CellAddress(maxColumn, maxRow).ToId(),
                    startSheet
                );
        return true;
    }

    // A DynamicRange evaluated as a scalar mirrors OFFSET/CHOOSE's range convention: it yields
    // ComputedValue.Reference(...) so a consumer that expects a reference (SUM, COUNT, ...) can still expand
    // it, degrading to #REF! only when the endpoints fail to resolve to a concrete range.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TryResolveReference(context, out var reference) ? ComputedValue.Reference(reference!)
        : HasScalarNameEndpoint(context) ? ComputedValue.Error(Error.Value)
        : ComputedValue.Error(Error.Ref);

    // A defined name used as a ':' endpoint can resolve to a scalar rather than a reference. The deleted-
    // reference parser preserves that name as an endpoint; Excel reports the malformed scalar range as
    // #VALUE!, while genuinely unresolvable endpoints retain the #REF! fallback.
    private bool HasScalarNameEndpoint(EvaluationContext context) =>
        IsScalarName(Start, context) || IsScalarName(End, context);

    private static bool IsScalarName(Expression endpoint, EvaluationContext context) =>
        endpoint is NameReference name
        && !name.TryResolveReference(context, out _)
        && name.Evaluate(context).Kind != ComputedValueKind.Error;

    // The sheet a resolved endpoint lives on. TryBox only accepts a cell, a rectangle or a zero-row
    // rectangle, so those are the only shapes reaching here.
    private static string SheetOf(Reference reference) =>
        reference switch
        {
            CellReference cell => cell.SheetName,
            EmptyRangeReference empty => empty.SheetName,
            _ => ((RangeReference)reference).SheetName,
        };

    private readonly record struct Box(int MinColumn, int MinRow, int MaxColumn, int MaxRow);

    // Bounding box of a resolved reference. Only bounded references (cell/rectangle) are supported as a
    // dynamic endpoint; an open reference as a dynamic endpoint is out of scope (returns false -> #REF!).
    private static bool TryBox(Reference reference, out Box box)
    {
        if (reference is CellReference cell)
        {
            var a = CellAddress.Parse(cell.Id);
            box = new Box(a.Column, a.Row, a.Column, a.Row);
            return true;
        }

        // A rectangle's normalized bounds (RangeBounds.TryFrom, the same min/max corners). A zero-row rectangle
        // (sweep item 33) ends one row ABOVE its anchor, so a span of two of them stays zero rows high while a
        // span reaching a real cell grows to include it.
        if (RangeBounds.TryFrom(reference, out var bounds))
        {
            box = new Box(bounds.LeftColumn, bounds.TopRow, bounds.RightColumn, bounds.BottomRow);
            return true;
        }

        box = default;
        return false;
    }
}
