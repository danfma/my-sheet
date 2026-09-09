namespace Danfma.MySheet.Expressions;

/// <summary>
/// Excel's implicit intersection (the "@" operator) applied at the CELL boundary: a formula whose FINAL
/// value is a multi-cell reference is not an error — it is intersected with the formula cell's own row and
/// column. A single-COLUMN reference spanning the formula's row yields that row's cell, a single-ROW
/// reference spanning the formula's column yields that column's cell, a 1x1 reference yields itself, and
/// anything wider on both axes is <c>#VALUE!</c> (Microsoft, "Implicit intersection operator: @").
///
/// <para>Only <see cref="Workbook.EvaluateCell"/> calls this, and only over the value
/// <see cref="NamedReferences.CaptureValue"/> captured for the cell's whole expression — INSIDE a formula a
/// range keeps being a range (<c>SUM(A1:A3)</c> is untouched), because there the consumer, not the cell,
/// decides what a multi-cell reference means. This is the RANGE half of the rule; the ARRAY half (a computed
/// array collapses to its top-left value) has no producer in this engine yet and is deliberately absent.</para>
///
/// <para>A <see cref="UnionReference"/> keeps its <c>#VALUE!</c>: a comma union of areas has no single row/
/// column axis to intersect against, so there is nothing to reason from.</para>
///
/// <para>Cross-sheet is POSITIONAL and sheet-independent: <c>=Sheet1!A1:A3</c> typed in <c>Sheet2!C2</c>
/// yields <c>Sheet1!A2</c>, because only the formula cell's row/column number enters the rule and the
/// dereference then happens on the reference's OWN sheet. That specific case was reasoned from the @
/// operator's definition (it is stated purely in terms of row and column numbers, with no sheet term), NOT
/// measured in Excel — the one claim here that is inference rather than observation.</para>
/// </summary>
internal static class ImplicitIntersection
{
    /// <summary>
    /// The value a cell shows for a reference-kind result: the intersection of <paramref name="reference"/>
    /// with the row and column of the cell being evaluated (<see cref="EvaluationContext.CellId"/>).
    /// </summary>
    public static ComputedValue Apply(Reference reference, EvaluationContext context)
    {
        // The intersection axis IS the formula cell's address, so a cell that has none — the rare non-A1 key
        // a host may store under, which Workbook.GetCellValueOverflow serves — has nothing to intersect with
        // and keeps the pre-existing #VALUE!.
        if (
            context.CellId is not { } cellId
            || !CellAddress.TryGetColumnRow(cellId, out var formulaColumn, out var formulaRow)
        )
        {
            return ComputedValue.Error(Error.Value);
        }

        switch (reference)
        {
            // A single cell is already the one value the boundary wants. No built-in producer reaches here
            // (INDIRECT and OFFSET dereference their 1x1 result themselves), but ComputedValue.Reference is
            // public API, so a host custom function can hand one back as a cell's result.
            case CellReference cell:
                return cell.Evaluate(context);

            case RangeReference range:
                // One parse of both corners for the whole intersection (GetBounds is the normalizing,
                // no-alloc idiom; its per-corner properties would re-parse).
                var bounds = range.GetBounds();

                return Intersect(
                    context,
                    range.SheetName,
                    bounds.LeftColumn,
                    bounds.RightColumn,
                    bounds.TopRow,
                    bounds.BottomRow,
                    formulaColumn,
                    formulaRow
                );

            // DECLARED bounds, deliberately NOT OpenRangeReference.ToBoundedRange's POPULATED box: in Excel
            // =A:A on row 7 is A7 even when A7 is empty (the boundary's blank→0 rule then shows 0), whereas
            // the populated box of a column filled only down to row 3 would answer #VALUE!. An unbounded
            // side is null and passes the containment test, which is what makes A:A single-column with a
            // free row axis and 1:1 single-row with a free column axis.
            case OpenRangeReference open:
                return Intersect(
                    context,
                    open.SheetName,
                    open.ColMin,
                    open.ColMax,
                    open.RowMin,
                    open.RowMax,
                    formulaColumn,
                    formulaRow
                );

            default:
                return ComputedValue.Error(Error.Value);
        }
    }

    // The shape rule, over nullable bounds so one implementation serves both the closed rectangle (every
    // bound known) and the open shapes (a null bound = that side is unbounded). Most specific branch first:
    // 1x1 needs no containment test at all, then the two single-axis cases.
    private static ComputedValue Intersect(
        EvaluationContext context,
        string sheetName,
        int? left,
        int? right,
        int? top,
        int? bottom,
        int formulaColumn,
        int formulaRow
    )
    {
        // The one column/row the reference occupies, when it occupies exactly one — null when that axis
        // spans more than a single line (an unbounded side always does).
        var fixedColumn = left is { } leftColumn && right == leftColumn ? leftColumn : (int?)null;
        var fixedRow = top is { } topRow && bottom == topRow ? topRow : (int?)null;

        // 1x1: the reference already denotes exactly one cell, wherever the formula sits.
        if (fixedColumn is { } onlyColumn && fixedRow is { } onlyRow)
        {
            return Dereference(context, sheetName, onlyColumn, onlyRow);
        }

        // A single COLUMN spanning the formula's row: that row's cell in the column.
        if (fixedColumn is { } column && Contains(formulaRow, top, bottom))
        {
            return Dereference(context, sheetName, column, formulaRow);
        }

        // A single ROW spanning the formula's column: that column's cell in the row.
        if (fixedRow is { } row && Contains(formulaColumn, left, right))
        {
            return Dereference(context, sheetName, formulaColumn, row);
        }

        // Wider than one cell on BOTH axes: no single answer, so #VALUE! (Excel's rule for a 2-D range at
        // the boundary). Should a both-axes rule ever be adopted — the formula cell's own (column, row) when
        // the rectangle contains it — it slots in here as one more branch:
        //   if (Contains(formulaColumn, left, right) && Contains(formulaRow, top, bottom)) …
        return ComputedValue.Error(Error.Value);
    }

    // Containment on one axis, with a null bound meaning "unbounded on that side" (always satisfied).
    private static bool Contains(int position, int? min, int? max) =>
        (min is not { } lower || position >= lower) && (max is not { } upper || position <= upper);

    // The allocation-free deref RangeReference.CellComputedValueAt already uses: numeric address against a
    // once-resolved sheet handle. It carries the full cell semantics — memoization, the cycle guard (so a
    // range that intersects its own formula cell answers #REF!) and the missing-sheet → #REF! rule, which is
    // why a reference to a sheet that does not exist needs no guard of its own here.
    private static ComputedValue Dereference(
        EvaluationContext context,
        string sheetName,
        int column,
        int row
    )
    {
        var workbook = context.Workbook;

        return workbook.GetCellValueDense(
            workbook.ResolveDenseHandle(sheetName),
            sheetName,
            column,
            row
        );
    }
}
