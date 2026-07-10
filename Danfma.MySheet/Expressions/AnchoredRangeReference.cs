using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// G3 spike (node-delta shared formulas): the range-endpoint twin of <see cref="AnchoredCellReference"/> — a
/// bounded <c>A1:B2</c>-shaped range inside a shared-formula MASTER tree, with both endpoints kept as
/// (column, row, $-anchor) components instead of normalized ids, so a <see cref="SharedFormulaSlave"/> can
/// shift each endpoint independently at evaluation time (an endpoint's row/column shifts only when its own
/// anchor flag is not set — exactly <c>Parser.ShiftCellId</c>'s per-token semantics).
///
/// Only produced by the Parser's anchored mode when BOTH range endpoints are simple cell references (see
/// <c>Parser.ParseRange</c>); an open range, a union, or a reference-returning endpoint is left as the
/// ordinary <see cref="OpenRangeReference"/>/<see cref="DynamicRange"/>/<see cref="UnionReference"/> node,
/// which the loader's shape check (<c>AnchoredFormulaSupport</c>) treats as NOT anchored-safe — that whole
/// shared-formula group falls back to the legacy token-delta expansion instead of guessing.
/// </summary>
[MemoryPackable]
public sealed partial record AnchoredRangeReference(
    int StartColumn,
    int StartRow,
    bool StartColumnAbsolute,
    bool StartRowAbsolute,
    int EndColumn,
    int EndRow,
    bool EndColumnAbsolute,
    bool EndRowAbsolute,
    [property: InternStringFormatter] string SheetName
) : Reference
{
    // A range has no scalar value — same rule as RangeReference.Evaluate.
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Error(Error.Value);

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = ToRangeReference(context);
        return true;
    }

    /// <summary>
    /// Materializes the delta-applied endpoints as a real <see cref="RangeReference"/>, so every range
    /// consumer (aggregate functions, <see cref="RangeReference.Expand(EvaluationContext)"/>,
    /// <see cref="RangeReference.ExpandComputedValues"/>) gets IDENTICAL behavior to the pre-spike expanded
    /// tree — this spike does not special-case range iteration, only avoids per-slave re-parsing. One small
    /// transient allocation per evaluation (two formatted ids); ranges are not the fixture's dominant shape
    /// (the K1-synthetic groups are all scalar cell arithmetic), so this was not optimized further — flagged
    /// as a production follow-up in the spike report.
    /// </summary>
    internal RangeReference ToRangeReference(EvaluationContext context)
    {
        var startColumn = StartColumnAbsolute ? StartColumn : StartColumn + context.DeltaColumn;
        var startRow = StartRowAbsolute ? StartRow : StartRow + context.DeltaRow;
        var endColumn = EndColumnAbsolute ? EndColumn : EndColumn + context.DeltaColumn;
        var endRow = EndRowAbsolute ? EndRow : EndRow + context.DeltaRow;

        return new RangeReference(
            new CellAddress(startColumn, startRow).ToId(),
            new CellAddress(endColumn, endRow).ToId(),
            SheetName
        );
    }
}
