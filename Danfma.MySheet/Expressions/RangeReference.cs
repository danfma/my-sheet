using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record RangeReference(string StartId, string EndId, string SheetName)
    : Reference
{
    // A range has no scalar value: used outside a function that accepts ranges it is a #VALUE! error.
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Error(Error.Value);

    // Backwards-compatible entry point used by tests and external callers.
    public IEnumerable<Expression> Expand(Workbook workbook) =>
        Expand(new EvaluationContext(workbook));

    /// <summary>
    /// Enumerates the stored expression of every cell in the rectangle (blank cells included as
    /// <see cref="BlankValue"/>). Reversed corners (e.g. <c>B2:A1</c>) are normalized.
    /// </summary>
    public IEnumerable<Expression> Expand(EvaluationContext context)
    {
        // Two layers guard a missing sheet. LOCAL guarantee: this enumerator never throws — a missing sheet
        // yields nothing (yield break) instead of a KeyNotFoundException. STRUCTURAL resolution: every
        // reference-consuming function first runs ReferenceGuard.MissingSheet and returns #REF! before it ever
        // enumerates, so a missing sheet surfaces as #REF! (Excel parity) rather than a silently empty range.
        if (!context.Workbook.Sheets.TryGetValue(SheetName, out var sheet))
        {
            yield break;
        }

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

    /// <summary>Enumerates the memoized <see cref="ComputedValue"/> of every cell in the rectangle (via the
    /// workbook cache) — the allocation-free path used by range-consuming functions.</summary>
    internal IEnumerable<ComputedValue> ExpandComputedValues(EvaluationContext context)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);

        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        // Resolve the sheet handle ONCE and address every cell numerically: a hit no longer builds an A1 id
        // (StringBuilder), re-parses it and re-resolves the sheet per cell — the memoization, cycle guard,
        // volatile taint and on-demand miss evaluation are unchanged (they live behind GetCellValueDense).
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(SheetName);

        for (var column = minColumn; column <= maxColumn; column++)
        {
            for (var row = minRow; row <= maxRow; row++)
            {
                yield return workbook.GetCellValueDense(handle, SheetName, column, row);
            }
        }
    }

    public int RowCount =>
        Math.Abs(CellAddress.Parse(EndId).Row - CellAddress.Parse(StartId).Row) + 1;

    public int ColumnCount =>
        Math.Abs(CellAddress.Parse(EndId).Column - CellAddress.Parse(StartId).Column) + 1;

    public int TopRow => Math.Min(CellAddress.Parse(StartId).Row, CellAddress.Parse(EndId).Row);

    public int LeftColumn =>
        Math.Min(CellAddress.Parse(StartId).Column, CellAddress.Parse(EndId).Column);

    /// <summary>The memoized <see cref="ComputedValue"/> at a 1-based (row, column) position (normalized corners).</summary>
    internal ComputedValue CellComputedValueAt(EvaluationContext context, int row, int column)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);
        var absoluteColumn = Math.Min(start.Column, end.Column) + column - 1;
        var absoluteRow = Math.Min(start.Row, end.Row) + row - 1;

        // Numeric address (no ToId round trip on a hit); same GetCellValue semantics via the dense twin.
        var workbook = context.Workbook;
        return workbook.GetCellValueDense(
            workbook.ResolveDenseHandle(SheetName),
            SheetName,
            absoluteColumn,
            absoluteRow
        );
    }
}
