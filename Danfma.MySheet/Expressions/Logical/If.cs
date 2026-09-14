using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record If(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // The condition slot accepts the TEXT "TRUE"/"FALSE" (case-insensitive, not trimmed) on top of the
        // usual truthiness — see ValueCoercion.CoerceToBoolAllowingTextWords. AND/OR/XOR must NOT: they
        // IGNORE text, so they keep the plain CoerceToBool through LogicalReduction.
        if (
            context
                .EvaluateConditionOnce(Arguments[0])
                .CoerceToBoolAllowingTextWords(out var condition) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        // Only the taken branch is computed (short-circuit), matching Excel.
        if (condition)
        {
            return ArrayEvaluation.BranchValue(Arguments[1], context);
        }

        return Arguments.Length == 3
            ? ArrayEvaluation.BranchValue(Arguments[2], context)
            : ComputedValue.Boolean(false);
    }

    // The value the taken branch hands back (ArrayEvaluation.BranchValue, shared with the mini-CSE's
    // scalar-condition build): a BARE-REFERENCE branch carries the reference it resolves to — sweep item
    // 32, the "IF returns a reference" decision for the scalar-condition shape, the same reference value a
    // range-bound NAME already flowed out of Evaluate. Range-aware consumers expand its cells
    // (SUM(IF(TRUE,A1:A3,0)) 14), the criteria family reads the range (COUNTIF 2), the cell boundary
    // intersects it (oracle 26.6.0, 2026-09-11: 9 at C3, "x" at C2 — plain entry), a single-cell branch
    // reaches the referenced-cell rule (text skipped: SUM 0), and anything else evaluates exactly as
    // before. An ARRAY-condition IF does not pass through here in selector role: its element-wise zip is
    // TryBuildIf's IfOperand, and its branches are values there, not references.
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        // The reference-reading twin of Evaluate — the CHOOSE twin is Choose.TryResolveReference — for the
        // consumers that RESOLVE their argument (ROWS/COLUMNS/INDEX/ISREF/AREAS, the criteria family's
        // fallback): the taken bare-reference branch's rectangle, or "not a reference" for anything else
        // (a scalar branch, a branch-less FALSE), exactly what the node's Evaluate would have carried.
        if (
            context
                .EvaluateConditionOnce(Arguments[0])
                .CoerceToBoolAllowingTextWords(out var condition) is
            { } error
        )
        {
            return false;
        }

        if (!condition)
        {
            return Arguments.Length == 3
                && Arguments[2].TryResolveReference(context, out reference);
        }

        return Arguments[1].TryResolveReference(context, out reference);
    }
}
