using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// G3 spike (node-delta shared formulas): the per-cell wrapper a shared-formula SLAVE stores instead of its
/// own fully-expanded tree. <see cref="Master"/> is ONE shared instance re-used by every slave in the group
/// (parsed once via the Parser's anchored mode, see <c>ExpressionParser.ParseAnchoredMasterBody</c>); only
/// <see cref="DeltaRow"/>/<see cref="DeltaColumn"/> — two ints — differ per slave cell. This is the entire
/// point of the spike: ~360k slave trees collapse to ~360k small wrapper nodes sharing a handful of master
/// trees (one per shared-formula group), instead of ~360k independent fully-expanded trees.
///
/// <see cref="Evaluate"/> pushes the delta into a fresh <see cref="EvaluationContext"/> and evaluates
/// <see cref="Master"/> against it; <see cref="AnchoredCellReference"/>/<see cref="AnchoredRangeReference"/>
/// nodes inside <see cref="Master"/> read that ambient delta to compute their effective address. The delta
/// does NOT leak into sub-cell evaluations — <see cref="Workbook.EvaluateCell"/> always mints a brand-new
/// <see cref="EvaluationContext"/> (delta defaults to 0) for every cell it evaluates, so a
/// <see cref="AnchoredCellReference"/> resolving to another cell (<c>Workbook.GetCellValueDense</c>) crosses
/// the cell boundary the same way <see cref="CellReference"/> always has; the delta only ever survives
/// WITHIN this one slave's own tree evaluation (LET/function-argument recursion), never across a
/// <c>GetCellValue</c>/<c>GetCellValueDense</c> call.
/// </summary>
[MemoryPackable]
public sealed partial record SharedFormulaSlave(Expression Master, int DeltaRow, int DeltaColumn) : Expression
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Master.Evaluate(context.WithDelta(DeltaRow, DeltaColumn));

    // A shared-formula slave is a whole-cell replacement for the plain expanded tree; a reference-context
    // consumer (the ':' operator, etc.) resolving THIS node directly should see through to what the delta-
    // shifted master resolves to, mirroring how an expanded tree would have behaved.
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference) =>
        Master.TryResolveReference(context.WithDelta(DeltaRow, DeltaColumn), out reference);
}
