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
    public override object? Compute(EvaluationContext context)
    {
        var value = Operand.Compute(context);

        if (ValueCoercion.TryToNumber(value, out var number) is { } error)
        {
            return error;
        }

        return Operator switch
        {
            UnaryOperator.Negate => -number,
            UnaryOperator.Plus => number,
            UnaryOperator.Percent => number / 100,
            _ => throw new ArgumentOutOfRangeException(nameof(Operator), Operator, null),
        };
    }
}
