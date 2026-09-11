using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Parsing;

/// <summary>
/// G3 spike (node-delta shared formulas): decides whether an anchored-mode master tree (built by
/// <see cref="ExpressionParser.ParseAnchoredMasterBody"/>) is safe to share across every slave of a group via
/// <see cref="SharedFormulaSlave"/>, or whether the group must fall back to the legacy per-slave token-delta
/// expansion (<see cref="ExpressionParser.ParseSharedFormulaBody"/>).
///
/// "Honest fallback" (per the spike design): the anchored/delta model only covers the minimum the shape needs
/// to be UNAMBIGUOUSLY correct — plain cell references (<see cref="AnchoredCellReference"/>), bounded ranges
/// (<see cref="AnchoredRangeReference"/>), arithmetic operators, defined names and structured table references
/// (both position-independent, so safe unshifted), and built-in/custom function calls whose arguments are
/// themselves supported. Anything
/// this Parser mode could not turn into an anchored node — an <see cref="OpenRangeReference"/> (whole column/
/// row), a <see cref="DynamicRange"/> (a reference-returning endpoint), a <see cref="UnionReference"/> (comma
/// union of areas) — is treated as UNSUPPORTED: rather than guess at their delta semantics, the whole group
/// reparses per-slave exactly as before the spike.
/// </summary>
internal static class AnchoredFormulaSupport
{
    /// <summary>
    /// True when <paramref name="expression"/> contains only node shapes the anchored/delta evaluation model
    /// can represent exactly.
    /// </summary>
    public static bool IsFullyAnchored(Expression expression) =>
        expression switch
        {
            AnchoredCellReference or AnchoredRangeReference => true,

            // A literal never depends on position.
            NumberValue or StringValue or BooleanValue or BlankValue or ErrorValue => true,

            // A defined name is resolved by name against Workbook.DefinedNames, and a structured reference by
            // table name against Workbook.Tables — both workbook-scoped registries — independent of the shared-
            // formula group's per-slave position, so both are safe to leave un-anchored (identical for every
            // slave, exactly as it is identical for every cell of an ordinary formula referencing the same name
            // or table). A TableReference carries no position component to shift: [@Col], the only position-
            // dependent structured form, is out of scope and has no TableArea member, and the delta a slave
            // pushes changes nothing in the node's resolution (StructuredReferenceSharedFormulaTests). For the
            // same reason both Parser modes, anchored and legacy-delta, must build the IDENTICAL node once the
            // arm exists (Phase 4 T5): a BracketedSpecifier token has no ($, column, row) component for either
            // mode to treat differently. Accepting rather than rejecting keeps the commonest real shape — one
            // table formula shared down thousands of rows — on this one-master-tree path instead of the
            // per-slave token re-parse (ExpressionParser.ParseSharedFormulaBody).
            // "Position-independent" is about the RESOLVED REFERENCE being delta-invariant, not about the
            // resulting VALUE: for a group of =SomeName or =Tabela1[Valor] cells where the reference denotes a
            // multi-cell range, each slave still shows a DIFFERENT value, because the cell boundary's implicit
            // intersection runs per cell from that cell's own CellId (Workbook.EvaluateCell) — measured on
            // Aspose.Cells 26.6.0 (2026-09-11, PLAIN and a SetSharedFormula group alike) as 10 / 20 / 30 down
            // the three data rows and #VALUE! on a row outside them, which is what the test above pins. That
            // leaves this verdict correct — nothing but the parsed tree is shared across the group.
            NameReference or TableReference => true,

            BinaryOperation binary => IsFullyAnchored(binary.Left) && IsFullyAnchored(binary.Right),

            UnaryOperation unary => IsFullyAnchored(unary.Operand),

            // A plain (non-anchored) CellReference/RangeReference should never appear inside a tree the
            // anchored Parser mode produced — BuildCellReference/ParseRange only emit those when NOT
            // anchored. Reject defensively rather than silently trust an un-anchored reference's literal id.
            CellReference or RangeReference => false,

            // Open ranges, dynamic (reference-returning-endpoint) ranges and comma-unions of areas are the
            // documented fallback triggers — the anchored mode does not attempt to model their per-slave
            // shift semantics.
            OpenRangeReference or DynamicRange or UnionReference => false,

            // Any built-in/custom function: supported when every argument is. Arguments come from the same
            // generic (name, args) accessor FormulaWriter/DependencyExtractor use, so this covers every
            // registered function uniformly (IF, ROUND, MAX, … and a user-named FunctionCall) without a
            // per-function whitelist; a function whose argument list this accessor cannot resolve is
            // conservatively rejected.
            Function function => TryGetArguments(function, out var arguments)
                && Array.TrueForAll(arguments, IsFullyAnchored),

            // Anything else (LET, array-only constructs, a future node type) — conservative reject.
            _ => false,
        };

    private static bool TryGetArguments(Function function, out Expression[] arguments)
    {
        try
        {
            (_, arguments) = FormulaWriter.Call(function);
            return true;
        }
        catch (NotSupportedException)
        {
            arguments = [];
            return false;
        }
    }
}
