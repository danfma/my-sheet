namespace Danfma.MySheet.Expressions;

/// <summary>
/// An EMPTY reference: a rectangle with ZERO rows on <see cref="SheetName"/>, anchored at
/// <see cref="TopRow"/> and spanning <see cref="LeftColumn"/>..<see cref="RightColumn"/> — what a structured
/// reference over a band with no rows resolves to (sweep item 33: a header-only table's <c>[#Data]</c>, whose
/// anchor is the row after the header). Measured on Aspose.Cells 26.6.0 (2026-09-14, both entry modes, the
/// order-independent reading the controller ruled binding): <c>SUM</c>/<c>COUNT</c>/<c>COUNTA</c> 0,
/// <c>ROWS</c> 0, <c>COLUMNS</c> the band's width, <c>AREAS</c> 1, <c>ISREF</c> TRUE, <c>ROW</c> the anchor
/// row, <c>AVERAGE</c> <c>#DIV/0!</c>, <c>INDEX(…,1,1)</c> <c>#REF!</c>, the lookups <c>#N/A</c>, the criteria
/// family 0, <c>FILTER</c> <c>#CALC!</c>.
/// <para>
/// A RUNTIME VALUE only: <see cref="TableReference"/> (and whatever resolves through it — a bare table name,
/// <c>INDIRECT</c>, a <c>':'</c> range between two such bands) hands it out as a resolved reference or a
/// <see cref="ComputedValue.Reference"/>. It is never parsed, never stored in a cell's tree and never
/// serialized, which is why it is deliberately NOT a <c>[MemoryPackUnion]</c> member of
/// <see cref="Expression"/>: what survives <c>Save</c>/<c>Load</c> is the table and the formula, and they
/// resolve to it again. It is not a <see cref="RangeReference"/> either, and cannot be: that node's corners
/// are min/max-normalized, so the inverted pair a zero-row band has would silently read the header row.
/// </para>
/// <para>
/// Consumers pattern-match it beside <see cref="RangeReference"/>: a value reader yields no cell
/// (<see cref="ComputedValue.EnumerateValues(EvaluationContext)"/>), geometry reads the anchor and the width,
/// the mini-CSE builds an array with no row (<c>ArrayEvaluation.BuildEmpty</c>), and a resize reads from the
/// anchor like any top-left (<c>OFFSET(T[Valor],0,0,1,1)</c> is the cell under the header).
/// </para>
/// </summary>
internal sealed record EmptyRangeReference(
    string SheetName,
    int TopRow,
    int LeftColumn,
    int RightColumn
) : Reference
{
    public int RowCount => 0;

    public int ColumnCount => RightColumn - LeftColumn + 1;

    /// <summary>
    /// The rectangle as <see cref="RangeBounds"/>: the real columns and a bottom row one ABOVE the anchor, so
    /// <see cref="RangeBounds.RowCount"/> is 0 by the bounds' own arithmetic and a loop over the rows never
    /// runs.
    /// </summary>
    internal RangeBounds GetBounds() => new(LeftColumn, TopRow, RightColumn, TopRow - 1);

    // A range has no scalar value: the same #VALUE! RangeReference.Evaluate answers.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Error(Error.Value);
}
