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
        if (Operand.Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Operator switch
        {
            UnaryOperator.Negate => ComputedValue.Number(-number),
            UnaryOperator.Plus => ComputedValue.Number(number),
            UnaryOperator.Percent => ComputedValue.Number(number / 100),
            _ => throw new ArgumentOutOfRangeException(nameof(Operator), Operator, null),
        };
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
