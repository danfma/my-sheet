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
/// (<see cref="AnchoredRangeReference"/>), arithmetic operators, defined names (position-independent, so
/// safe unshifted), and built-in/custom function calls whose arguments are themselves supported. Anything
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

            // A defined name is resolved by name against Workbook.DefinedNames, independent of the shared-
            // formula group's per-slave position — safe to leave un-anchored (identical for every slave,
            // exactly as it is identical for every cell of an ordinary formula referencing the same name).
            NameReference => true,

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
