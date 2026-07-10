using MemoryPack;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// G3 spike (node-delta shared formulas): a cell reference INSIDE a shared-formula master tree that keeps
/// its Excel <c>$</c> anchors instead of discarding them (unlike <see cref="CellReference"/>, whose
/// <c>Id</c> is normalized by <c>Parser.NormalizeCellId</c> with the anchors stripped — the reason the
/// AST-shift approach was rejected for the per-slave shifter). <see cref="Column"/>/<see cref="Row"/> are the
/// MASTER's own 1-based coordinates; the effective cell a given slave reads is resolved at evaluation time
/// from <see cref="EvaluationContext.DeltaRow"/>/<see cref="EvaluationContext.DeltaColumn"/> — the axis stays
/// put when the matching <c>*Absolute</c> flag is set (an anchored <c>$A$1</c> reads the same cell from every
/// slave), otherwise it shifts by the ambient delta (a relative <c>A1</c> reads a different cell per slave),
/// exactly mirroring <c>Parser.ShiftCellId</c>'s per-slave text-shift semantics but without re-parsing.
///
/// Only ever produced by the Parser's ANCHORED mode (see <c>ExpressionParser.ParseAnchoredMasterBody</c>),
/// used exclusively to parse a shared-formula group's MASTER once; it is not emitted by normal formula
/// parsing and never appears outside a <see cref="SharedFormulaSlave"/>'s <see cref="SharedFormulaSlave.Master"/>
/// tree (which is why it is safe for <see cref="EvaluationContext.DeltaRow"/> to default to 0 everywhere
/// else — a bare <see cref="AnchoredCellReference"/> evaluated with a zero-delta context degrades exactly to
/// its own literal (Column, Row), which is also correct for the master cell itself when it is rendered
/// through the same anchored tree in <see cref="Parsing.FormulaWriter"/>).
/// </summary>
[MemoryPackable]
public sealed partial record AnchoredCellReference(
    int Column,
    int Row,
    bool ColumnAbsolute,
    bool RowAbsolute,
    [property: InternStringFormatter] string SheetName
) : Reference
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var (column, row) = Effective(context);
        var workbook = context.Workbook;

        // Numeric-address fast path (Workbook.GetCellValueDense): identical memoization, cycle-guard and
        // volatile-taint semantics as the string-keyed CellReference.Evaluate, minus the per-evaluation A1
        // id string this node would otherwise have to format (there is no pre-shifted string to reuse — that
        // is the whole point of NOT expanding a tree per slave). A cache MISS still materializes the id once,
        // exactly like GetCellValueDense's existing miss path.
        return workbook.GetCellValueDense(workbook.ResolveDenseHandle(SheetName), SheetName, column, row);
    }

    /// <summary>
    /// Resolves to a concrete, delta-applied <see cref="CellReference"/> — used by reference-context
    /// consumers (the ':' range operator combining a computed endpoint, etc.). NOTE: this does NOT make an
    /// anchored node interchangeable with a real <see cref="CellReference"/> for consumers that pattern-match
    /// the concrete type directly (e.g. ROW/COLUMN/ADDRESS special-case <c>is CellReference</c>);
    /// see the spike report for that documented production gap.
    /// </summary>
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        var (column, row) = Effective(context);
        reference = new CellReference(new CellAddress(column, row).ToId(), SheetName);
        return true;
    }

    internal (int Column, int Row) Effective(EvaluationContext context) =>
        (
            ColumnAbsolute ? Column : Column + context.DeltaColumn,
            RowAbsolute ? Row : Row + context.DeltaRow
        );
}
