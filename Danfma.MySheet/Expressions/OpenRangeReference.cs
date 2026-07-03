using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A reference that is unbounded on at least one side — a whole column (<c>A:A</c>, <c>A:C</c>), a whole
/// row (<c>1:1</c>, <c>1:5</c>) or a one-sided open range (<c>A2:A</c>, <c>A:A10</c>, <c>A1:C</c>). Each
/// limit is <c>null</c> when that side is open on that axis: the LEFT endpoint gives the lower bounds
/// (<see cref="ColMin"/>/<see cref="RowMin"/>), the RIGHT endpoint the upper bounds
/// (<see cref="ColMax"/>/<see cref="RowMax"/>). A column-only endpoint informs no row and a row-only
/// endpoint no column, so that axis stays open on that side. (When all four limits are known the parser
/// produces a plain <see cref="RangeReference"/> instead.)
///
/// <para>Semantics are "the POPULATED cells within the limits": enumeration scans <c>Sheet.Cells</c> and
/// keeps the cells whose (column,row) fall inside the non-null bounds (a null bound always passes). Blank
/// cells contribute 0, so <c>SUM(A:A)</c> matches Excel while never materializing the empty grid.</para>
/// </summary>
[MemoryPackable]
public sealed partial record OpenRangeReference(
    int? ColMin,
    int? ColMax,
    int? RowMin,
    int? RowMax,
    string SheetName
) : Reference
{
    /// <summary>
    /// Builds a normalized open range: swaps reversed corners so <c>min ≤ max</c> on each axis where both
    /// limits are known (e.g. <c>C:A</c> ⇒ columns A..C). The parser uses this so the stored record is
    /// always normalized.
    /// </summary>
    public static OpenRangeReference Create(
        int? colMin,
        int? colMax,
        int? rowMin,
        int? rowMax,
        string sheetName
    )
    {
        if (colMin is { } cLow && colMax is { } cHigh && cLow > cHigh)
        {
            (colMin, colMax) = (colMax, colMin);
        }

        if (rowMin is { } rLow && rowMax is { } rHigh && rLow > rHigh)
        {
            (rowMin, rowMax) = (rowMax, rowMin);
        }

        return new OpenRangeReference(colMin, colMax, rowMin, rowMax, sheetName);
    }

    // A range has no scalar value: used outside a function that accepts ranges it is a #VALUE! error,
    // exactly like a bare RangeReference.
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Error(Error.Value);

    private bool Contains(int column, int row) =>
        (ColMin is not { } colMin || column >= colMin)
        && (ColMax is not { } colMax || column <= colMax)
        && (RowMin is not { } rowMin || row >= rowMin)
        && (RowMax is not { } rowMax || row <= rowMax);

    /// <summary>Backwards-compatible entry point mirroring <see cref="RangeReference.Expand(Workbook)"/>.</summary>
    public IEnumerable<Expression> Expand(Workbook workbook) => Expand(new EvaluationContext(workbook));

    /// <summary>
    /// The NaiveScan: yields the id of every POPULATED cell within the limits. Iterates
    /// <c>Sheet.Cells.Keys</c> and extracts (column,row) with the no-alloc
    /// <see cref="CellAddress.TryGetColumnRow"/> — no substring, no <c>int.Parse</c>.
    /// </summary>
    internal IEnumerable<string> PopulatedIds(EvaluationContext context)
    {
        var sheet = context.Workbook.Sheets[SheetName];

        foreach (var id in sheet.Keys)
        {
            if (CellAddress.TryGetColumnRow(id, out var column, out var row) && Contains(column, row))
            {
                yield return id;
            }
        }
    }

    /// <summary>Enumerates the stored expression of every POPULATED cell within the limits (NaiveScan).</summary>
    public IEnumerable<Expression> Expand(EvaluationContext context)
    {
        var sheet = context.Workbook.Sheets[SheetName];

        foreach (var id in PopulatedIds(context))
        {
            yield return sheet[id];
        }
    }

    /// <summary>
    /// The allocation-free <see cref="ComputedValue"/> view of the POPULATED cells within the limits — the
    /// NaiveScan the aggregate functions consume (memoized value per cell).
    /// </summary>
    internal IEnumerable<ComputedValue> ExpandComputedValues(EvaluationContext context)
    {
        foreach (var id in PopulatedIds(context))
        {
            yield return context.Workbook.GetCellValue(SheetName, id);
        }
    }

    /// <summary>
    /// Scans for the populated bounding box within the limits. Returns <c>false</c> when no populated cell
    /// falls inside the bounds (an empty selection).
    /// </summary>
    private bool TryGetPopulatedBounds(
        EvaluationContext context,
        out int minColumn,
        out int maxColumn,
        out int minRow,
        out int maxRow
    )
    {
        minColumn = int.MaxValue;
        maxColumn = int.MinValue;
        minRow = int.MaxValue;
        maxRow = int.MinValue;

        var any = false;
        var sheet = context.Workbook.Sheets[SheetName];

        foreach (var id in sheet.Keys)
        {
            if (!CellAddress.TryGetColumnRow(id, out var column, out var row) || !Contains(column, row))
            {
                continue;
            }

            any = true;
            if (column < minColumn) minColumn = column;
            if (column > maxColumn) maxColumn = column;
            if (row < minRow) minRow = row;
            if (row > maxRow) maxRow = row;
        }

        return any;
    }

    /// <summary>
    /// Resolves this open range to the concrete <see cref="RangeReference"/> of its POPULATED bounding box,
    /// used where a bounded range is required (VLOOKUP/HLOOKUP table, INDEX, OFFSET base). Returns
    /// <c>null</c> when the selection is empty (no populated cell within the limits).
    /// </summary>
    internal RangeReference? ToBoundedRange(EvaluationContext context) =>
        TryGetPopulatedBounds(context, out var minColumn, out var maxColumn, out var minRow, out var maxRow)
            ? new RangeReference(
                new CellAddress(minColumn, minRow).ToId(),
                new CellAddress(maxColumn, maxRow).ToId(),
                SheetName
            )
            : null;

    /// <summary>
    /// The row extent for <c>ROWS</c>: a bounded row axis is structural (<c>max − min + 1</c>); an open row
    /// axis uses the POPULATED extent within the limits (0 when empty). This DIVERGES from Excel, which
    /// reports the fixed grid height (1,048,576) — a gridless model has no such grid.
    /// </summary>
    internal int RowExtent(EvaluationContext context) =>
        RowMin is { } min && RowMax is { } max
            ? max - min + 1
            : TryGetPopulatedBounds(context, out _, out _, out var minRow, out var maxRow)
                ? maxRow - minRow + 1
                : 0;

    /// <summary>
    /// The column extent for <c>COLUMNS</c>: a bounded column axis is structural and exact
    /// (<c>COLUMNS(A:C)</c> = 3); an open column axis uses the POPULATED extent within the limits (0 when
    /// empty).
    /// </summary>
    internal int ColumnExtent(EvaluationContext context) =>
        ColMin is { } min && ColMax is { } max
            ? max - min + 1
            : TryGetPopulatedBounds(context, out var minColumn, out var maxColumn, out _, out _)
                ? maxColumn - minColumn + 1
                : 0;
}
