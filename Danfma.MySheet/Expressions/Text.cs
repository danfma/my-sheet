using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Text(Expression[] Arguments) : Function
{
    // TEXT(value, format) — delegates to .NET numeric formatting, which covers the common Excel codes
    // (0, #, ".", thousands "," and "%"). Date/text/colour codes are out of scope.
    public override object? Compute(EvaluationContext context)
    {
        if (
            ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var number) is
            { } numberError
        )
        {
            return numberError;
        }

        if (
            ValueCoercion.TryToText(Arguments[1].Compute(context), out var format) is
            { } formatError
        )
        {
            return formatError;
        }

        try
        {
            return ExcelDateFormat.IsDateOrTime(format)
                ? DateTime
                    .FromOADate(number)
                    .ToString(ExcelDateFormat.ToDotNet(format), CultureInfo.InvariantCulture)
                : number.ToString(format, CultureInfo.InvariantCulture);
        }
        catch (Exception exception)
            when (exception is FormatException or ArgumentException or OverflowException)
        {
            return ErrorValue.NotValue;
        }
    }
}
