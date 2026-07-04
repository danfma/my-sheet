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
// SheetName carries the same read-side interning as CellReference (wire byte-identical; see that file).
[MemoryPackable]
public sealed partial record OpenRangeReference(
    int? ColMin,
    int? ColMax,
    int? RowMin,
    int? RowMax,
    [property: InternStringFormatter] string SheetName
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

    /// <summary>Backwards-compatible entry point mirroring <see cref="RangeReference.Expand(Workbook)"/>.</summary>
    public IEnumerable<Expression> Expand(Workbook workbook) => Expand(new EvaluationContext(workbook));

    /// <summary>
    /// Yields the <c>(column, row)</c> of every POPULATED cell within the limits, in column-then-row order,
    /// using the Layer-1 <see cref="SheetStructuralIndex"/> so a narrow reference in a big sheet no longer scans
    /// every key — it visits only the columns (or, for a whole row, the rows) the reference covers. The index
    /// stores the coordinates numerically, so this path parses no id at all.
    /// </summary>
    internal IEnumerable<(int Column, int Row)> PopulatedCells(EvaluationContext context)
    {
        // Two layers guard a missing sheet. LOCAL guarantee: this scan never throws — a missing sheet yields
        // nothing (yield break) instead of a KeyNotFoundException. STRUCTURAL resolution: every reference-
        // consuming function first runs ReferenceGuard.MissingSheet and returns #REF! before it ever
        // enumerates, so a missing sheet surfaces as #REF! (Excel parity) — never a silently empty column.
        if (!context.Workbook.Sheets.TryGetValue(SheetName, out var sheet))
        {
            yield break;
        }

        // The structural index is lifetime-scoped and write-maintained (3.0): built once on the sheet's first
        // open-range read and kept current by SetCell/Remove thereafter, so it is always available here.
        var index = sheet.GetStructuralIndex();

        // Whole-row reference (column axis fully open, row axis bounded): drive the symmetric ROW index so
        // a whole-row read touches only its rows, never the big columns. Yields row-then-column order; with
        // no column bound to apply, every column in the row's (column-sorted) list qualifies.
        if (ColMin is null && ColMax is null && RowMin is { } wholeRowMin && RowMax is { } wholeRowMax)
        {
            for (var row = wholeRowMin; row <= wholeRowMax; row++)
            {
                if (index.TryGetRow(row, out var rowColumns))
                {
                    foreach (var column in rowColumns)
                    {
                        yield return (column, row);
                    }
                }
            }

            yield break;
        }

        // Column-driven path (every other shape): visit only the covered columns, each list already row-
        // sorted. A row bound narrows each list to its qualifying slice by binary search — no per-id
        // parsing, and a pure whole column (A:A) yields its list verbatim.
        foreach (var column in ColumnsToVisit(index))
        {
            if (!index.TryGetColumn(column, out var columnRows))
            {
                continue;
            }

            var first = RowMin is { } rowMin ? FirstAtOrAboveRow(columnRows, rowMin) : 0;
            var last = RowMax is { } rowMax ? LastAtOrBelowRow(columnRows, rowMax) : columnRows.Count - 1;

            for (var i = first; i <= last; i++)
            {
                yield return (column, columnRows[i]);
            }
        }
    }

    /// <summary>
    /// The id of every POPULATED cell within the limits, in column-then-row order. A COLD-path convenience over
    /// <see cref="PopulatedCells"/>: it re-derives each A1 id from the numeric coordinates. The hot value path
    /// (<see cref="ExpandComputedValues"/>) consumes the numeric pairs directly and never materializes an id.
    /// </summary>
    internal IEnumerable<string> PopulatedIds(EvaluationContext context)
    {
        foreach (var (column, row) in PopulatedCells(context))
        {
            yield return new CellAddress(column, row).ToId();
        }
    }

    // The columns the reference covers, ascending. When both column limits are known it is the finite
    // inclusive range [ColMin, ColMax] (the fast, common case: A:A, A:C, A2:A, A:A10, A1:C). When a column
    // side is open it is the populated columns within whatever bound IS known, sorted so enumeration stays
    // column-ascending.
    private IEnumerable<int> ColumnsToVisit(SheetStructuralIndex index)
    {
        if (ColMin is { } lowerColumn && ColMax is { } upperColumn)
        {
            for (var column = lowerColumn; column <= upperColumn; column++)
            {
                yield return column;
            }

            yield break;
        }

        var matching = new List<int>();

        foreach (var column in index.ColumnKeys)
        {
            if ((ColMin is not { } min || column >= min) && (ColMax is not { } max || column <= max))
            {
                matching.Add(column);
            }
        }

        matching.Sort();

        foreach (var column in matching)
        {
            yield return column;
        }
    }

    /// <summary>Enumerates the stored expression of every POPULATED cell within the limits. A COLD path: it
    /// re-derives each id to look the expression up by key.</summary>
    public IEnumerable<Expression> Expand(EvaluationContext context)
    {
        if (!context.Workbook.Sheets.TryGetValue(SheetName, out var sheet))
        {
            yield break;
        }

        foreach (var (column, row) in PopulatedCells(context))
        {
            yield return sheet[new CellAddress(column, row).ToId()];
        }
    }

    /// <summary>
    /// The allocation-free <see cref="ComputedValue"/> view of the POPULATED cells within the limits — the
    /// index-backed enumeration the aggregate functions consume (memoized value per cell).
    /// </summary>
    internal IEnumerable<ComputedValue> ExpandComputedValues(EvaluationContext context)
    {
        // The structural index yields (column, row) pairs directly (3.3: it stores coordinates numerically, not
        // id strings), so this hot path is fully parse-free: it feeds the dense accessor the numeric address with
        // a hoisted sheet handle — no per-cell id materialization, no re-parse, no per-cell sheet-handle lookup.
        // GetCellValueDense re-derives an id ONLY on a cache MISS, purely to locate the expression to evaluate.
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(SheetName);

        foreach (var (column, row) in PopulatedCells(context))
        {
            yield return workbook.GetCellValueDense(handle, SheetName, column, row);
        }
    }

    /// <summary>
    /// Computes the populated bounding box within the limits. Returns <c>false</c> when no populated cell
    /// falls inside the bounds (an empty selection, or a missing sheet — no throw, exactly as before).
    /// Because each index list is sorted (a column's ids by row, a row's ids by column), the box comes from
    /// the FIRST and LAST qualifying entry of each covered list — O(covered lists · log n), never a walk
    /// over every id. VLOOKUP/INDEX resolve their open table through here on every evaluation, so this must
    /// not re-enumerate the whole column.
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

        if (!context.Workbook.Sheets.TryGetValue(SheetName, out var sheet))
        {
            return false;
        }

        var index = sheet.GetStructuralIndex();

        // Whole-row shape: each covered row's list is column-sorted, so its first/last ids give the column
        // extremes directly (no column bound applies on this branch, by definition).
        if (ColMin is null && ColMax is null && RowMin is { } wholeRowMin && RowMax is { } wholeRowMax)
        {
            for (var row = wholeRowMin; row <= wholeRowMax; row++)
            {
                if (!index.TryGetRow(row, out var rowColumns) || rowColumns.Count == 0)
                {
                    continue;
                }

                any = true;
                if (row < minRow) minRow = row;
                if (row > maxRow) maxRow = row;

                // The list is column-sorted, so its ends ARE the column extremes.
                var firstColumn = rowColumns[0];
                var lastColumn = rowColumns[^1];
                if (firstColumn < minColumn) minColumn = firstColumn;
                if (lastColumn > maxColumn) maxColumn = lastColumn;
            }

            return any;
        }

        // Column-driven shapes: each covered column's list is row-sorted, so the row extremes within the
        // row bounds are found by binary search at both ends.
        foreach (var column in ColumnsToVisit(index))
        {
            if (!index.TryGetColumn(column, out var columnRows) || columnRows.Count == 0)
            {
                continue;
            }

            var first = RowMin is { } rowMin ? FirstAtOrAboveRow(columnRows, rowMin) : 0;
            var last = RowMax is { } rowMax ? LastAtOrBelowRow(columnRows, rowMax) : columnRows.Count - 1;

            if (first > last)
            {
                continue; // no populated cell of this column falls inside the row bounds
            }

            any = true;
            if (column < minColumn) minColumn = column;
            if (column > maxColumn) maxColumn = column;

            var lowRow = columnRows[first];
            var highRow = columnRows[last];
            if (lowRow < minRow) minRow = lowRow;
            if (highRow > maxRow) maxRow = highRow;
        }

        return any;
    }

    // Binary search over a row-sorted list: the index of the first entry whose row is >= target
    // (list.Count when none).
    private static int FirstAtOrAboveRow(List<int> sortedRows, int target)
    {
        var low = 0;
        var high = sortedRows.Count;

        while (low < high)
        {
            var mid = (low + high) >> 1;

            if (sortedRows[mid] < target)
            {
                low = mid + 1;
            }
            else
            {
                high = mid;
            }
        }

        return low;
    }

    // The index of the last entry whose row is <= target (-1 when none).
    private static int LastAtOrBelowRow(List<int> sortedRows, int target) =>
        FirstAtOrAboveRow(sortedRows, target + 1) - 1;

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
