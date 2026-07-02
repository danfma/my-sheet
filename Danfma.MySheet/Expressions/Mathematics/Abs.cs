using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record Abs(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Abs(number));
    }
}
