using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Round(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var number) is { } numberError)
        {
            return numberError;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(context), out var digits) is { } digitsError)
        {
            return digitsError;
        }

        var factor = Math.Pow(10, digits);

        return Math.Round(number * factor, MidpointRounding.AwayFromZero) / factor;
    }
}
