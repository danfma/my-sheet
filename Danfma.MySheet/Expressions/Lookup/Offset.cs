using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Offset(Expression[] Arguments) : Function
{
    // OFFSET(reference, rows, cols, [height], [width]) — scalar (single-cell) result only for now.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (TryComputeTarget(context, out var target) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return target is CellReference cell
            ? context.Workbook.GetCellValue(cell.SheetName, cell.Id) // 1x1: dereference
            : ComputedValue.Reference(target!); // multi-cell: reference value
    }

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
        => TryComputeTarget(context, out reference) is null;

    // Computes OFFSET's target reference (a CellReference for 1x1, else a RangeReference) without
    // dereferencing. Returns null on success (reference set to the target); otherwise the exact
    // error that should propagate to the caller (reference left null).
    private Error? TryComputeTarget(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        // The base may be written directly or through a defined name that stands for a cell/range.
        if (
            !NamedReferences.TryResolveReference(Arguments[0], context, out var baseReference)
            || !TryBase(baseReference, out var sheetName, out var baseColumn, out var baseRow)
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

        var height = 1.0;
        var width = 1.0;

        if (Arguments.Length >= 4 && Arguments[3].Evaluate(context).CoerceToNumber(out height) is { } e1)
        {
            return e1;
        }

        if (Arguments.Length >= 5 && Arguments[4].Evaluate(context).CoerceToNumber(out width) is { } e2)
        {
            return e2;
        }

        var startColumn = baseColumn + (int)columns;
        var startRow = baseRow + (int)rows;

        if (startColumn < 1 || startRow < 1)
        {
            return Error.Ref;
        }

        reference =
            height == 1 && width == 1
                ? new CellReference(new CellAddress(startColumn, startRow).ToId(), sheetName)
                : new RangeReference(
                    new CellAddress(startColumn, startRow).ToId(),
                    new CellAddress(startColumn + (int)width - 1, startRow + (int)height - 1).ToId(),
                    sheetName
                );

        return null;
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
