using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Text(Expression[] Arguments) : Function
{
    // TEXT(value, format) — delegates to .NET numeric formatting, which covers the common Excel codes
    // (0, #, ".", thousands "," and "%"). Date/text/colour codes are out of scope.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        if (Arguments[1].Evaluate(context).CoerceToText(out var format) is { } formatError)
        {
            return ComputedValue.Error(formatError);
        }

        try
        {
            return ComputedValue.Text(
                ExcelDateFormat.IsDateOrTime(format)
                    ? DateTime
                        .FromOADate(number)
                        .ToString(ExcelDateFormat.ToDotNet(format), CultureInfo.InvariantCulture)
                    : number.ToString(format, CultureInfo.InvariantCulture)
            );
        }
        catch (Exception exception)
            when (exception is FormatException or ArgumentException or OverflowException)
        {
            return ComputedValue.Error(Error.Value);
        }
    }
}
