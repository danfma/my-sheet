using MemoryPack;

namespace MySheet.Expressions;

public enum UnaryOperator
{
    // NOTE: append-only — MemoryPack serializes enum members by their underlying value.
    Negate,
    Plus,
}

[MemoryPackable]
public sealed partial record UnaryOperation(UnaryOperator Operator, Expression Operand) : Expression
{
    public override object? Compute(Workbook workbook)
    {
        var value = Operand.Compute(workbook);

        if (ValueCoercion.TryToNumber(value, out var number) is { } error)
        {
            return error;
        }

        return Operator switch
        {
            UnaryOperator.Negate => -number,
            UnaryOperator.Plus => number,
            _ => throw new ArgumentOutOfRangeException(nameof(Operator), Operator, null),
        };
    }
}
