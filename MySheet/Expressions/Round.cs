using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Round(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(workbook), out var number) is { } numberError)
        {
            return numberError;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(workbook), out var digits) is { } digitsError)
        {
            return digitsError;
        }

        var factor = Math.Pow(10, digits);

        return Math.Round(number * factor, MidpointRounding.AwayFromZero) / factor;
    }
}
