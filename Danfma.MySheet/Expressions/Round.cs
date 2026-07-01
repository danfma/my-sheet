using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Round(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var digits) is { } digitsError)
        {
            return ComputedValue.Error(digitsError);
        }

        var factor = Math.Pow(10, digits);

        return ComputedValue.Number(Math.Round(number * factor, MidpointRounding.AwayFromZero) / factor);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
