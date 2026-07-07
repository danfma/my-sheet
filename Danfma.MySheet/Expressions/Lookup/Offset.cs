using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Offset(Expression[] Arguments) : Function
{
    // OFFSET(reference, rows, cols, [height], [width]) — scalar (single-cell) result only for now.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            TryComputeTarget(
                context,
                out var sheetName,
                out var startColumn,
                out var startRow,
                out var height,
                out var width
            ) is { } error
        )
        {
            return ComputedValue.Error(error);
        }

        // 1x1: dereference directly, no CellReference allocation (matches the original).
        if (height == 1 && width == 1)
        {
            return context.Workbook.GetCellValue(sheetName, new CellAddress(startColumn, startRow).ToId());
        }

        return ComputedValue.Reference(BuildRange(sheetName, startColumn, startRow, height, width));
    }

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        if (
            TryComputeTarget(
                context,
                out var sheetName,
                out var startColumn,
                out var startRow,
                out var height,
                out var width
            ) is not null
        )
        {
            reference = null;
            return false;
        }

        reference =
            height == 1 && width == 1
                ? new CellReference(new CellAddress(startColumn, startRow).ToId(), sheetName)
                : BuildRange(sheetName, startColumn, startRow, height, width);

        return true;
    }

    // Computes OFFSET's target geometry (sheet + top-left cell + height/width) without allocating a
    // reference. Returns null on success (out parameters set to the target); otherwise the exact
    // error that should propagate to the caller.
    private Error? TryComputeTarget(
        EvaluationContext context,
        out string sheetName,
        out int startColumn,
        out int startRow,
        out int height,
        out int width
    )
    {
        sheetName = string.Empty;
        startColumn = 0;
        startRow = 0;
        height = 1;
        width = 1;

        // The base may be written directly or through a defined name that stands for a cell/range.
        if (
            !NamedReferences.TryResolveReference(Arguments[0], context, out var baseReference)
            || !TryBase(baseReference, out sheetName, out var baseColumn, out var baseRow)
        )
        {
            return Error.Ref;
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var rows) is { } rowsError)
        {
            return rowsError;
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var columns) is { } columnsError)
        {
            return columnsError;
        }

        var h = 1.0;
        var w = 1.0;

        if (Arguments.Length >= 4 && Arguments[3].Evaluate(context).CoerceToNumber(out h) is { } e1)
        {
            return e1;
        }

        if (Arguments.Length >= 5 && Arguments[4].Evaluate(context).CoerceToNumber(out w) is { } e2)
        {
            return e2;
        }

        startColumn = baseColumn + (int)columns;
        startRow = baseRow + (int)rows;

        if (startColumn < 1 || startRow < 1)
        {
            return Error.Ref;
        }

        height = (int)h;
        width = (int)w;

        return null;
    }

    private static RangeReference BuildRange(string sheetName, int startColumn, int startRow, int height, int width) =>
        new(
            new CellAddress(startColumn, startRow).ToId(),
            new CellAddress(startColumn + width - 1, startRow + height - 1).ToId(),
            sheetName
        );

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
