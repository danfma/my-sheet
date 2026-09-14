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
            ) is
            { } error
        )
        {
            return error;
        }

        // 1x1: dereference directly, no CellReference allocation (matches the original).
        if (height == 1 && width == 1)
        {
            return context.Workbook.GetCellValue(
                sheetName,
                new CellAddress(startColumn, startRow).ToId()
            );
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
            )
            is not null
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
    // error RESULT that should propagate to the caller — as a ComputedValue so the base-reference arm
    // can hand back the node's own error (sweep item 34(b)).
    private ComputedValue? TryComputeTarget(
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

        // The base may be written directly or through a defined name that stands for a cell/range. When
        // it does not resolve, the node's OWN error is the answer (sweep item 34(b): #NAME? for an
        // unknown name, the node's #REF! for an unresolvable structured reference) — OFFSET's own #REF!
        // stays only for a base that is merely not a cell/range, and for a target pushed off the grid.
        //
        // boundOpenRanges: false — sweep item 37 follow-up, ruling (a): an open base ($5:$1000, A:A, …)
        // stays open here, so TryBase's own OpenRangeReference arm reads its ABSOLUTE first cell (column A
        // / row 1 when that side is open) instead of the POPULATED bounding box's own corner — the same
        // fix INDEX's open-base form needed, for the same reason (Index.IndexIntoOpenRange's remarks).
        if (
            !NamedReferences.TryResolveReference(
                Arguments[0],
                context,
                out var baseReference,
                boundOpenRanges: false
            )
        )
        {
            return ReferencePosition.Unresolved(
                Arguments[0],
                context,
                ComputedValue.Error(Error.Ref)
            );
        }

        if (!TryBase(baseReference, out sheetName, out var baseColumn, out var baseRow))
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

        // Excel truncates height/width toward zero (like rows/columns), so OFFSET(A1,0,0,1.9) is a
        // single row, not two. (int) truncates toward zero, matching (int)rows / (int)columns above.
        if (Arguments.Length >= 4)
        {
            if (Arguments[3].Evaluate(context).CoerceToNumber(out var h) is { } e1)
            {
                return ComputedValue.Error(e1);
            }

            height = (int)h;
        }

        if (Arguments.Length >= 5)
        {
            if (Arguments[4].Evaluate(context).CoerceToNumber(out var w) is { } e2)
            {
                return ComputedValue.Error(e2);
            }

            width = (int)w;
        }

        startColumn = baseColumn + (int)columns;
        startRow = baseRow + (int)rows;

        if (startColumn < 1 || startRow < 1)
        {
            return ComputedValue.Error(Error.Ref);
        }

        // A non-positive size is #REF! in Excel (height/width of 0, or a fraction that truncates to 0).
        // This also prevents building an invalid cell id like "A0" from a zero-height/width range. (Excel's
        // negative height/width extends in the opposite direction; that abs+direction case is not modeled
        // here — a rare form left as #REF! rather than a wrong value.)
        if (height < 1 || width < 1)
        {
            return ComputedValue.Error(Error.Ref);
        }

        return null;
    }

    private static RangeReference BuildRange(
        string sheetName,
        int startColumn,
        int startRow,
        int height,
        int width
    ) =>
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

            // Sweep item 33: a zero-row rectangle's top-left is its anchor, the row after a header-only
            // table's header — OFFSET(Tabela1[Valor],0,0,1,1) reads that cell (oracle: 7 with a 7 there).
            case EmptyRangeReference empty:
                sheetName = empty.SheetName;
                column = empty.LeftColumn;
                row = empty.TopRow;
                return true;

            // Sweep item 37 follow-up, ruling (a): an open base's "first cell" is its ABSOLUTE (row 1,
            // column A) corner when a side is open, or its declared bound when it is not — the same
            // AbsoluteRow(1)/AbsoluteColumn(1) INDEX's own open-base form uses.
            case OpenRangeReference open:
                sheetName = open.SheetName;
                column = open.AbsoluteColumn(1);
                row = open.AbsoluteRow(1);
                return true;

            default:
                sheetName = string.Empty;
                column = 0;
                row = 0;
                return false;
        }
    }
}
