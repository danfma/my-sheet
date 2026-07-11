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
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Error(Error.Value);

    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = ToRangeReference(context);
        return true;
    }

    /// <summary>
    /// Materializes the delta-applied endpoints as a real <see cref="RangeReference"/>, for a COLD consumer
    /// that genuinely needs a syntactic reference node (e.g. <see cref="TryResolveReference"/>, the ':' range
    /// operator, <see cref="ArrayEvaluation"/>'s once-per-call setup). One small transient allocation (two
    /// formatted ids) — fine off the hot per-cell path; <see cref="ExpandComputedValues"/> below is the
    /// allocation-free twin for the site that IS hot.
    /// </summary>
    internal RangeReference ToRangeReference(EvaluationContext context)
    {
        var (startColumn, startRow, endColumn, endRow) = EffectiveEndpoints(
            context.DeltaRow,
            context.DeltaColumn
        );

        return new RangeReference(
            new CellAddress(startColumn, startRow).ToId(),
            new CellAddress(endColumn, endRow).ToId(),
            SheetName
        );
    }

    /// <summary>
    /// The delta-applied endpoints as raw ints (NOT corner-normalized — mirrors the exact pre-existing
    /// <see cref="ToRangeReference"/> computation, only pulled out so a caller without an
    /// <see cref="EvaluationContext"/> can share it). <see cref="DirtyGraph.DependencyExtractor"/> uses this
    /// to build a <see cref="DirtyGraph.RangeDep"/> (which normalizes itself via <c>Math.Min</c>/<c>Math.Max</c>,
    /// exactly like the plain <see cref="RangeReference"/> case does) without materializing a transient
    /// <see cref="RangeReference"/> + two formatted ids.
    /// </summary>
    internal (int StartColumn, int StartRow, int EndColumn, int EndRow) EffectiveEndpoints(
        int deltaRow,
        int deltaColumn
    ) =>
        (
            StartColumnAbsolute ? StartColumn : StartColumn + deltaColumn,
            StartRowAbsolute ? StartRow : StartRow + deltaRow,
            EndColumnAbsolute ? EndColumn : EndColumn + deltaColumn,
            EndRowAbsolute ? EndRow : EndRow + deltaRow
        );

    /// <summary>
    /// The normalized rectangle bounds (min/max corners) after the delta — the <see cref="RangeReference.GetBounds"/>
    /// idiom, computed directly from <see cref="EffectiveEndpoints"/> with no intermediate <see cref="RangeReference"/>.
    /// </summary>
    internal RangeBounds GetBounds(EvaluationContext context)
    {
        var (startColumn, startRow, endColumn, endRow) = EffectiveEndpoints(
            context.DeltaRow,
            context.DeltaColumn
        );

        return new RangeBounds(
            Math.Min(startColumn, endColumn),
            Math.Min(startRow, endRow),
            Math.Max(startColumn, endColumn),
            Math.Max(startRow, endRow)
        );
    }

    /// <summary>
    /// The allocation-free view of the delta-applied range's memoized values — the exact twin of
    /// <see cref="RangeReference.ExpandComputedValues"/>, but reading bounds straight from
    /// <see cref="GetBounds(EvaluationContext)"/> instead of routing through a transient
    /// <see cref="ToRangeReference"/> first. This is the HOT site: <see cref="NumericAggregation"/>'s
    /// per-argument case for this node type is exercised once per aggregate-function evaluation, and a
    /// shared-formula range group (e.g. the K1-synthetic fixture's <c>SUM(B2:D2)</c> group, 60k slaves) calls
    /// it once per slave cell — the two formatted-id strings <see cref="ToRangeReference"/> would have
    /// allocated per evaluation are gone.
    /// </summary>
    internal RangeValueSequence ExpandComputedValues(EvaluationContext context)
    {
        var bounds = GetBounds(context);

        var workbook = context.Workbook;
        var store = workbook.DenseStore;
        var handle = store.HandleFor(SheetName);

        return new RangeValueSequence(
            workbook,
            store,
            SheetName,
            handle,
            bounds.LeftColumn,
            bounds.RightColumn,
            bounds.TopRow,
            bounds.BottomRow
        );
    }
}
