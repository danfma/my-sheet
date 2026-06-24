using MemoryPack;

namespace MySheet.Expressions;

public enum BinaryOperator
{
    // NOTE: append-only — MemoryPack serializes enum members by their underlying value.
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Equal,
    NotEqual,
    LessThan,
    GreaterThan,
    LessThanOrEqual,
    GreaterThanOrEqual,
}

[MemoryPackable]
public sealed partial record BinaryOperation(BinaryOperator Operator, Expression Left, Expression Right) : Expression
{
    public override object? Compute(Workbook workbook)
    {
        var leftValue = Left.Compute(workbook);
        var rightValue = Right.Compute(workbook);

        if (ValueCoercion.TryToNumber(leftValue, out var left) is { } leftError)
        {
            return leftError;
        }

        if (ValueCoercion.TryToNumber(rightValue, out var right) is { } rightError)
        {
            return rightError;
        }

        return Operator switch
        {
            BinaryOperator.Add => left + right,
            BinaryOperator.Subtract => left - right,
            BinaryOperator.Multiply => left * right,
            BinaryOperator.Divide => right == 0 ? ErrorValue.DivByZero : left / right,
            BinaryOperator.Power => Math.Pow(left, right),
            BinaryOperator.Equal => left == right,
            BinaryOperator.NotEqual => left != right,
            BinaryOperator.LessThan => left < right,
            BinaryOperator.GreaterThan => left > right,
            BinaryOperator.LessThanOrEqual => left <= right,
            BinaryOperator.GreaterThanOrEqual => left >= right,
            _ => throw new ArgumentOutOfRangeException(nameof(Operator), Operator, null),
        };
    }
}
