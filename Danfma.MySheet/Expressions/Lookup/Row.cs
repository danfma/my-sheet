using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Row(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        // A reference to a missing sheet is a structural #REF!, not a row position.
        ReferenceGuard.MissingSheet(Arguments, context)
            is { } missing
            ? ComputedValue.Error(missing)
            : Arguments switch
            {
                [CellReference cell] => ComputedValue.Number(CellAddress.Parse(cell.Id).Row),
                [RangeReference range] => ComputedValue.Number(range.TopRow),
                // Phase 2 audit (shared-formula delta production): mirrors the two cases above for a
                // ROW(ref) argument written INSIDE a shared-formula master — the argument is an anchored
                // node (its own $-anchors preserved, delta applied at evaluation time), not a plain
                // CellReference/RangeReference, so it fell to `_ => #VALUE!` before this fix (a slave's
                // ROW(A1) silently broke while ROW() with no argument — using context.CellId, see below —
                // was already correct by construction).
                [AnchoredCellReference anchoredCell] => ComputedValue.Number(
                    anchoredCell.Effective(context).Row
                ),
                [AnchoredRangeReference anchoredRange] => ComputedValue.Number(
                    anchoredRange.ToRangeReference(context).TopRow
                ),
                // ROW() with no argument uses the cell currently being evaluated, when one is known. Inside a
                // SharedFormulaSlave this is the SLAVE's own cell (EvaluationContext.WithDelta leaves
                // SheetName/CellId untouched — only the ambient delta changes), so this arm is already
                // correct for a shared-formula slave without any change; see SharedFormulaSlaveFunctionTests.
                [] when context.CellId is { } id => ComputedValue.Number(CellAddress.Parse(id).Row),
                _ => ComputedValue.Error(Error.Value),
            };
}
