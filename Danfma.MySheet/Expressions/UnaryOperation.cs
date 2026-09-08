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
        // (issue #8). A range/union NODE is captured as a reference VALUE (CaptureValue) rather than
        // evaluated — evaluating it yields #VALUE! (a range has no scalar) — so value-path consumers
        // (SUM(+A1:A3), MATCH(x, +A1:C1, 0)) still see the cells; syntactic consumers (ROWS, INDEX) reach
        // the node through TryResolveReference below.
        if (Operator == UnaryOperator.Plus)
        {
            var value = NamedReferences.CaptureValue(Operand, context);

            return value.Kind == ComputedValueKind.Blank ? ComputedValue.Number(0) : value;
        }

        if (Operand.Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Operator switch
        {
            UnaryOperator.Negate => ComputedValue.Number(-number),
            UnaryOperator.Percent => ComputedValue.Number(number / 100),
            _ => throw new ArgumentOutOfRangeException(nameof(Operator), Operator, null),
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
