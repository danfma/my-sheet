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
        var operand = Operand.Evaluate(context);

        // Unary '+' is Excel's legacy (Lotus) no-op: the operand comes back unchanged, TYPE included — text
        // stays text, TRUE stays a boolean, a reference stays a reference for range consumers (SUM(+A1:A3)),
        // an error propagates. Only a blank becomes 0. Coercing to a number here (as '-' and '%' do) turned
        // `=+A1` on a text cell into #VALUE! and TRUE into 1 (issue #8).
        if (Operator == UnaryOperator.Plus)
        {
            return operand.Kind == ComputedValueKind.Blank ? ComputedValue.Number(0) : operand;
        }

        if (operand.CoerceToNumber(out var number) is { } error)
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
}
