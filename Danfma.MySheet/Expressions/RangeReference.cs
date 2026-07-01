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

        for (var column = minColumn; column <= maxColumn; column++)
        {
            for (var row = minRow; row <= maxRow; row++)
            {
                yield return context.Workbook.GetCellComputedValue(
                    SheetName,
                    new CellAddress(column, row).ToId()
                );
            }
        }
    }

    /// <summary>Boxed (<c>object?</c>) view of <see cref="ExpandComputedValues"/>, for interop.</summary>
    public IEnumerable<object?> ExpandValues(EvaluationContext context)
    {
        foreach (var value in ExpandComputedValues(context))
        {
            yield return value.AsObject();
        }
    }

    public int RowCount =>
        Math.Abs(CellAddress.Parse(EndId).Row - CellAddress.Parse(StartId).Row) + 1;

    public int ColumnCount =>
        Math.Abs(CellAddress.Parse(EndId).Column - CellAddress.Parse(StartId).Column) + 1;

    public int TopRow => Math.Min(CellAddress.Parse(StartId).Row, CellAddress.Parse(EndId).Row);

    /// <summary>The memoized <see cref="ComputedValue"/> at a 1-based (row, column) position (normalized corners).</summary>
    internal ComputedValue CellComputedValueAt(EvaluationContext context, int row, int column)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);
        var id = new CellAddress(
            Math.Min(start.Column, end.Column) + column - 1,
            Math.Min(start.Row, end.Row) + row - 1
        ).ToId();

        return context.Workbook.GetCellComputedValue(SheetName, id);
    }

    /// <summary>Boxed (<c>object?</c>) view of <see cref="CellComputedValueAt"/>, for interop.</summary>
    public object? CellValueAt(EvaluationContext context, int row, int column) =>
        CellComputedValueAt(context, row, column).AsObject();
}
