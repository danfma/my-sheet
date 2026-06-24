using MemoryPack;

namespace Danfma.MySheet.Expressions;

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
    Concat,
}

[MemoryPackable]
public sealed partial record BinaryOperation(
    BinaryOperator Operator,
    Expression Left,
    Expression Right
) : Expression
{
    public override object? Compute(EvaluationContext context)
    {
        var leftValue = Left.Compute(context);
        if (leftValue is ErrorValue leftError)
        {
            return leftError;
        }

        var rightValue = Right.Compute(context);
        if (rightValue is ErrorValue rightError)
        {
            return rightError;
        }

        // Equality and ordering compare across types (Excel order: number < text < boolean); only the
        // arithmetic operators coerce to a number.
        switch (Operator)
        {
            case BinaryOperator.Equal:
                return ValueCoercion.AreEqual(leftValue, rightValue);
            case BinaryOperator.NotEqual:
                return !ValueCoercion.AreEqual(leftValue, rightValue);
            case BinaryOperator.LessThan:
                return ValueCoercion.Compare(leftValue, rightValue) < 0;
            case BinaryOperator.GreaterThan:
                return ValueCoercion.Compare(leftValue, rightValue) > 0;
            case BinaryOperator.LessThanOrEqual:
                return ValueCoercion.Compare(leftValue, rightValue) <= 0;
            case BinaryOperator.GreaterThanOrEqual:
                return ValueCoercion.Compare(leftValue, rightValue) >= 0;

            case BinaryOperator.Concat:
                if (ValueCoercion.TryToText(leftValue, out var leftText) is { } leftTextError)
                {
                    return leftTextError;
                }

                if (ValueCoercion.TryToText(rightValue, out var rightText) is { } rightTextError)
                {
                    return rightTextError;
                }

                return leftText + rightText;
        }

        if (ValueCoercion.TryToNumber(leftValue, out var left) is { } leftNumberError)
        {
            return leftNumberError;
        }

        if (ValueCoercion.TryToNumber(rightValue, out var right) is { } rightNumberError)
        {
            return rightNumberError;
        }

        return Operator switch
        {
            BinaryOperator.Add => left + right,
            BinaryOperator.Subtract => left - right,
            BinaryOperator.Multiply => left * right,
            BinaryOperator.Divide => right == 0 ? ErrorValue.DivByZero : left / right,
            BinaryOperator.Power => Math.Pow(left, right),
            _ => throw new ArgumentOutOfRangeException(nameof(Operator), Operator, null),
        };
    }
}
