using MemoryPack;

namespace Danfma.MySheet.Expressions;

public enum UnaryOperator
{
    // NOTE: append-only — MemoryPack serializes enum members by their underlying value.
    Negate,
    Plus,
    Percent,
}

[MemoryPackable]
public sealed partial record UnaryOperation(UnaryOperator Operator, Expression Operand) : Expression
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // Unary '+' is Excel's legacy (Lotus) no-op: the operand comes back unchanged, TYPE included — text
        // stays text, TRUE stays a boolean, an error propagates. Only a blank becomes 0. Coercing to a
        // number here (as '-' and '%' do) turned `=+A1` on a text cell into #VALUE! and TRUE into 1
        // (issue #8). The operand is a BINDING SITE (Phase 11c, ArrayBindings.Capture): a range/union NODE
        // is captured as a reference VALUE rather than evaluated — evaluating it yields #VALUE! (a range has
        // no scalar) — so value-path consumers (SUM(+A1:A3), MATCH(x, +A1:C1, 0)) still see the cells and
        // syntactic consumers (ROWS, INDEX) reach the node through TryResolveReference below; a computed
        // ARRAY operand is built once and read as its TOP-LEFT, the @ rule a bare producer follows
        // (=+FILTER(A1:A3,A1:A3>0) is 5, =+(A1:A3*2) is 10 — Aspose.Cells 26.6.0, 2026-09-11, H20, the
        // second in the array-entered column where plain entry answers #VALUE!). A CONSUMER of the same node
        // streams the whole array instead: '+' is transparent to the mini-CSE (ArrayEvaluation.TryBuildUnary).
        if (Operator == UnaryOperator.Plus)
        {
            var value = ArrayBindings.Capture(Operand, context).TopLeft;

            return value.Kind == ComputedValueKind.Blank ? ComputedValue.Number(0) : value;
        }

        return Apply(Operator, Operand.Evaluate(context));
    }

    /// <summary>
    /// Applies <c>-</c> or <c>%</c> to an already-computed value, reusing the exact scalar semantics for both
    /// the normal <see cref="Evaluate"/> path and the element-wise array path (<c>UnaryOperand</c> in
    /// <see cref="ArrayEvaluation"/>): coerce to a number (an error, a non-numeric text, propagates as the
    /// coercion's error), then negate or divide by 100. Never <see cref="UnaryOperator.Plus"/> — both callers
    /// route that operator elsewhere (the reference-preserving path above; the mini-CSE's TRANSPARENT arm,
    /// which hands the operand's own array back unwrapped precisely because a <c>+</c> applies nothing per
    /// element), so the last arm is unreachable and THROWS. Answering <c>#VALUE!</c> there would turn a
    /// routing bug into a wrong cell value — an error indistinguishable from a legitimate coercion failure,
    /// silently broadcast over every element of an array — where an exception names the operator and the
    /// caller that mis-routed it. Same arm, same exception, as <c>BinaryOperation.Apply</c>.
    /// </summary>
    internal static ComputedValue Apply(UnaryOperator @operator, in ComputedValue operand)
    {
        if (operand.CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return @operator switch
        {
            UnaryOperator.Negate => ComputedValue.Number(-number),
            UnaryOperator.Percent => ComputedValue.Number(number / 100),
            _ => throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null),
        };
    }

    // The no-op is transparent to reference-context consumers too: ROWS(+A1:A3), INDEX(+A1:A3, 2) and the
    // ':' operator resolve through '+' to the operand's reference. '-' and '%' produce numbers, never references.
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        if (Operator == UnaryOperator.Plus)
        {
            return Operand.TryResolveReference(context, out reference);
        }

        reference = null;
        return false;
    }
}
