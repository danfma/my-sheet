using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Offset(Expression[] Arguments) : Function
{
    // OFFSET(reference, rows, cols, [height], [width]) — scalar (single-cell) result only for now.
    public override object? Compute(EvaluationContext context)
    {
        if (!TryBase(Arguments[0], out var sheetName, out var baseColumn, out var baseRow))
        {
            return ErrorValue.Reference;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(context), out var rows) is { } rowsError)
        {
            return rowsError;
        }

        if (ValueCoercion.TryToNumber(Arguments[2].Compute(context), out var columns) is { } columnsError)
        {
            return columnsError;
        }

        var height = 1.0;
        var width = 1.0;

        if (Arguments.Length >= 4 && ValueCoercion.TryToNumber(Arguments[3].Compute(context), out height) is { } e1)
        {
            return e1;
        }

        if (Arguments.Length >= 5 && ValueCoercion.TryToNumber(Arguments[4].Compute(context), out width) is { } e2)
        {
            return e2;
        }

        var startColumn = baseColumn + (int)columns;
        var startRow = baseRow + (int)rows;

        if (startColumn < 1 || startRow < 1)
        {
            return ErrorValue.Reference;
        }

        if (height == 1 && width == 1)
        {
            return context.Workbook.GetCellValue(sheetName, new CellAddress(startColumn, startRow).ToId());
        }

        // A multi-cell result is a range, expanded by functions that accept ranges (e.g. SUM(OFFSET(...))).
        return new RangeReference(
            new CellAddress(startColumn, startRow).ToId(),
            new CellAddress(startColumn + (int)width - 1, startRow + (int)height - 1).ToId(),
            sheetName);
    }

    private static bool TryBase(Expression reference, out string sheetName, out int column, out int row)
    {
        switch (reference)
        {
            case CellReference cell:
                var cellAddress = CellAddress.Parse(cell.Id);
                sheetName = cell.SheetName;
                column = cellAddress.Column;
                row = cellAddress.Row;
                return true;

            case RangeReference range:
                var start = CellAddress.Parse(range.StartId);
                sheetName = range.SheetName;
                column = start.Column;
                row = start.Row;
                return true;

            default:
                sheetName = string.Empty;
                column = 0;
                row = 0;
                return false;
        }
    }
}
