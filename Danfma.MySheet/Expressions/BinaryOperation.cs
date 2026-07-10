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
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var left = Left.Evaluate(context);
        if (left.Kind == ComputedValueKind.Error)
        {
            return left;
        }

        var right = Right.Evaluate(context);
        if (right.Kind == ComputedValueKind.Error)
        {
            return right;
        }

        return Apply(Operator, left, right);
    }

    /// <summary>
    /// Applies a binary operator to two already-computed values, reusing the exact scalar semantics for both
    /// the normal <see cref="Evaluate"/> path and the element-wise array path (<see cref="ArrayEvaluation"/>).
    /// An error on either side short-circuits and is returned verbatim (so an errored cell is preserved at its
    /// position in an array), mirroring the operand-error propagation of <see cref="Evaluate"/>.
    /// </summary>
    internal static ComputedValue Apply(
        BinaryOperator @operator,
        in ComputedValue left,
        in ComputedValue right
    )
    {
        if (left.Kind == ComputedValueKind.Error)
        {
            return left;
        }

        if (right.Kind == ComputedValueKind.Error)
        {
            return right;
        }

        // Equality and ordering compare across types (Excel order: number < text < boolean); only the
        // arithmetic operators coerce to a number.
        switch (@operator)
        {
            case BinaryOperator.Equal:
                return ComputedValue.Boolean(ValueCoercion.AreEqual(left, right));
            case BinaryOperator.NotEqual:
                return ComputedValue.Boolean(!ValueCoercion.AreEqual(left, right));
            case BinaryOperator.LessThan:
                return ComputedValue.Boolean(ValueCoercion.Compare(left, right) < 0);
            case BinaryOperator.GreaterThan:
                return ComputedValue.Boolean(ValueCoercion.Compare(left, right) > 0);
            case BinaryOperator.LessThanOrEqual:
                return ComputedValue.Boolean(ValueCoercion.Compare(left, right) <= 0);
            case BinaryOperator.GreaterThanOrEqual:
                return ComputedValue.Boolean(ValueCoercion.Compare(left, right) >= 0);

            case BinaryOperator.Concat:
                if (left.CoerceToText(out var leftText) is { } leftTextError)
                {
                    return ComputedValue.Error(leftTextError);
                }

                if (right.CoerceToText(out var rightText) is { } rightTextError)
                {
                    return ComputedValue.Error(rightTextError);
                }

                return ComputedValue.Text(leftText + rightText);
        }

        if (left.CoerceToNumber(out var l) is { } leftNumberError)
        {
            return ComputedValue.Error(leftNumberError);
        }

        if (right.CoerceToNumber(out var r) is { } rightNumberError)
        {
            return ComputedValue.Error(rightNumberError);
        }

        return @operator switch
        {
            BinaryOperator.Add => ComputedValue.Number(l + r),
            BinaryOperator.Subtract => ComputedValue.Number(l - r),
            BinaryOperator.Multiply => ComputedValue.Number(l * r),
            BinaryOperator.Divide => r == 0
                ? ComputedValue.Error(Error.DivZero)
                : ComputedValue.Number(l / r),
            BinaryOperator.Power => ComputedValue.Number(Math.Pow(l, r)),
            _ => throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null),
        };
    }
}
