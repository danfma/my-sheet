using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Offset(Expression[] Arguments) : Function
{
    // OFFSET(reference, rows, cols, [height], [width]) — scalar (single-cell) result only for now.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (!TryBase(Arguments[0], out var sheetName, out var baseColumn, out var baseRow))
        {
            return ComputedValue.Error(Error.Ref);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var rows) is { } rowsError)
        {
            return ComputedValue.Error(rowsError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var columns) is { } columnsError)
        {
            return ComputedValue.Error(columnsError);
        }

        var height = 1.0;
        var width = 1.0;

        if (Arguments.Length >= 4 && Arguments[3].Evaluate(context).CoerceToNumber(out height) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (Arguments.Length >= 5 && Arguments[4].Evaluate(context).CoerceToNumber(out width) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        var startColumn = baseColumn + (int)columns;
        var startRow = baseRow + (int)rows;

        if (startColumn < 1 || startRow < 1)
        {
            return ComputedValue.Error(Error.Ref);
        }

        if (height == 1 && width == 1)
        {
            return context.Workbook.GetCellComputedValue(
                sheetName,
                new CellAddress(startColumn, startRow).ToId()
            );
        }

        // A multi-cell result is a range (Reference kind), expanded by functions that accept ranges.
        return ComputedValue.Reference(
            new RangeReference(
                new CellAddress(startColumn, startRow).ToId(),
                new CellAddress(startColumn + (int)width - 1, startRow + (int)height - 1).ToId(),
                sheetName
            )
        );
    }

    private static bool TryBase(
        Expression reference,
        out string sheetName,
        out int column,
        out int row
    )
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
