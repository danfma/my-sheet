using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

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
                    // The unchecked map, not ToDateTime: an unrepresentable serial is #VALUE! here (the
                    // catch below), not the #NUM! the date functions answer.
                    ? DateSerial
                        .ToDateTimeUnchecked(number)
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
