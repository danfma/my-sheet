using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Building a serial from parts: DATE (calendar overflow like Excel), TIME (time-of-day fraction, mod 24h),
// and the text parsers DATEVALUE/TIMEVALUE. See DateSerial/DateTextParser for the shared machinery.

[MemoryPackable]
public sealed partial record Date(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var year) is { } yearError)
        {
            return ComputedValue.Error(yearError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var month) is { } monthError)
        {
            return ComputedValue.Error(monthError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var day) is { } dayError)
        {
            return ComputedValue.Error(dayError);
        }

        return DateSerial.FromComponents(year, month, day, out var serial) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(serial);
    }
}

[MemoryPackable]
public sealed partial record Time(Expression[] Arguments) : Function
{
    // Hour/minute/second are each 0..32767; any component negative or above 32767 → #NUM!. Overflow within
    // range rolls up (750 minutes → 12h30m) and the whole thing is taken mod 24h into a [0,1) fraction.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var hour) is { } hourError)
        {
            return ComputedValue.Error(hourError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var minute) is { } minuteError)
        {
            return ComputedValue.Error(minuteError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var second) is { } secondError)
        {
            return ComputedValue.Error(secondError);
        }

        var h = Math.Truncate(hour);
        var m = Math.Truncate(minute);
        var s = Math.Truncate(second);

        if (OutOfRange(h) || OutOfRange(m) || OutOfRange(s))
        {
            return ComputedValue.Error(Error.Num);
        }

        var totalSeconds = h * 3600d + m * 60d + s;
        var fraction = totalSeconds % 86400d / 86400d;

        return ComputedValue.Number(fraction);
    }

    private static bool OutOfRange(double component) => component < 0d || component > 32767d;
}

[MemoryPackable]
public sealed partial record DateValue(Expression[] Arguments) : Function
{
    // DATEVALUE(date_text) → whole-day serial (the time part of a date+time string is dropped). Unparseable
    // text → #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return DateTextParser.TryParseDate(text, out var date)
            ? ComputedValue.Number(Math.Floor(DateSerial.FromDateTime(date)))
            : ComputedValue.Error(Error.Value);
    }
}

[MemoryPackable]
public sealed partial record TimeValue(Expression[] Arguments) : Function
{
    // TIMEVALUE(time_text) → time-of-day fraction in [0,1) (any date part is ignored). Unparseable → #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return DateTextParser.TryParseTime(text, out var fraction)
            ? ComputedValue.Number(fraction)
            : ComputedValue.Error(Error.Value);
    }
}
