using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

[MemoryPackable]
public sealed partial record Text(Expression[] Arguments) : Function
{
    // TEXT(value, format) — date/time codes go through ExcelDateFormat (which prints Excel's own calendar,
    // day zero and the phantom 1900-02-29 included); everything else delegates to .NET numeric formatting,
    // which covers the common Excel codes (0, #, ".", thousands "," and "%"). Colour and section codes are
    // out of scope.
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

        if (ExcelDateFormat.IsDateOrTime(format))
        {
            // TEXT keeps its OWN range policy: a serial no date can represent is #VALUE! here (measured
            // TEXT(-1,…) and TEXT(2958466,…)), not the #NUM! the date functions answer — which is why
            // TryRender reports the range verdict as a bool instead of propagating an Error.
            return ExcelDateFormat.TryRender(format, number, out var rendered)
                ? ComputedValue.Text(rendered)
                : ComputedValue.Error(Error.Value);
        }

        try
        {
            return ComputedValue.Text(number.ToString(format, CultureInfo.InvariantCulture));
        }
        catch (Exception exception)
            when (exception is FormatException or ArgumentException or OverflowException)
        {
            return ComputedValue.Error(Error.Value);
        }
    }
}
