namespace Danfma.MySheet.Expressions;

internal static class ResolvedReferenceValue
{
    public static ComputedValue Read(
        Reference reference,
        EvaluationContext context,
        out bool absent
    )
    {
        if (TryGetSingletonCell(reference, out var cell))
        {
            absent =
                context.Workbook.Sheets.TryGetValue(cell.SheetName, out var sheet)
                && !sheet.ContainsKey(cell.Id);
            return cell.Evaluate(context);
        }

        absent = false;
        return ComputedValue.Reference(reference);
    }

    private static bool TryGetSingletonCell(Reference reference, out CellReference cell)
    {
        if (reference is CellReference direct)
        {
            cell = direct;
            return true;
        }

        if (
            reference is RangeReference range
            && RangeBounds.TryFrom(range, out var bounds)
            && bounds is { RowCount: 1, ColumnCount: 1 }
        )
        {
            cell = new CellReference(
                new CellAddress(bounds.LeftColumn, bounds.TopRow).ToId(),
                range.SheetName
            );
            return true;
        }

        cell = null!;
        return false;
    }
}
