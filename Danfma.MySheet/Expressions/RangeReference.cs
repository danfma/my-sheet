using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record RangeReference(string StartId, string EndId, string SheetName)
    : Reference
{
    // A range has no scalar value: used outside a function that accepts ranges it is a #VALUE! error.
    public override object? Compute(EvaluationContext context) => ErrorValue.NotValue;

    // Backwards-compatible entry point used by tests and external callers.
    public IEnumerable<Expression> Expand(Workbook workbook) =>
        Expand(new EvaluationContext(workbook));

    /// <summary>
    /// Enumerates the stored expression of every cell in the rectangle (blank cells included as
    /// <see cref="BlankValue"/>). Reversed corners (e.g. <c>B2:A1</c>) are normalized.
    /// </summary>
    public IEnumerable<Expression> Expand(EvaluationContext context)
    {
        var sheet = context.Workbook.Sheets[SheetName];
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);

        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        for (var column = minColumn; column <= maxColumn; column++)
        {
            for (var row = minRow; row <= maxRow; row++)
            {
                yield return sheet[new CellAddress(column, row).ToId()];
            }
        }
    }

    /// <summary>Enumerates the memoized value of every cell in the rectangle (via the workbook cache).</summary>
    public IEnumerable<object?> ExpandValues(EvaluationContext context)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);

        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        for (var column = minColumn; column <= maxColumn; column++)
        {
            for (var row = minRow; row <= maxRow; row++)
            {
                yield return context.Workbook.GetCellValue(
                    SheetName,
                    new CellAddress(column, row).ToId()
                );
            }
        }
    }

    public int RowCount =>
        Math.Abs(CellAddress.Parse(EndId).Row - CellAddress.Parse(StartId).Row) + 1;

    public int ColumnCount =>
        Math.Abs(CellAddress.Parse(EndId).Column - CellAddress.Parse(StartId).Column) + 1;

    public int TopRow => Math.Min(CellAddress.Parse(StartId).Row, CellAddress.Parse(EndId).Row);

    /// <summary>Returns the memoized value of the cell at a 1-based (row, column) position (normalized corners).</summary>
    public object? CellValueAt(EvaluationContext context, int row, int column)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);
        var id = new CellAddress(
            Math.Min(start.Column, end.Column) + column - 1,
            Math.Min(start.Row, end.Row) + row - 1
        ).ToId();

        return context.Workbook.GetCellValue(SheetName, id);
    }
}
