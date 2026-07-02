using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record RoundUp(Expression[] Arguments) : Function
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
        var scaled = number * factor;
        // Round away from zero.
        var rounded = scaled < 0 ? Math.Floor(scaled) : Math.Ceiling(scaled);

        return ComputedValue.Number(rounded / factor);
    }
}
