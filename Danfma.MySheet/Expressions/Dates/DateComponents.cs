using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Component extraction from a date serial: YEAR/MONTH/DAY read the calendar part via OADate; HOUR/MINUTE/
// SECOND read the time-of-day fraction (rounded to the nearest second, matching Excel). A negative serial
// is out of range → #NUM!. Text-of-date arguments are NOT accepted (only numeric serials / numeric text
// that CoerceToNumber already parses) — see the date namespace note.

[MemoryPackable]
public sealed partial record Year(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return DateSerial.ToDateTime(serial, out var date) is { } rangeError
            ? ComputedValue.Error(rangeError)
            : ComputedValue.Number(date.Year);
    }
}

[MemoryPackable]
public sealed partial record Month(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return DateSerial.ToDateTime(serial, out var date) is { } rangeError
            ? ComputedValue.Error(rangeError)
            : ComputedValue.Number(date.Month);
    }
}

[MemoryPackable]
public sealed partial record Day(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return DateSerial.ToDateTime(serial, out var date) is { } rangeError
            ? ComputedValue.Error(rangeError)
            : ComputedValue.Number(date.Day);
    }
}

[MemoryPackable]
public sealed partial record Hour(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (DateSerial.ToDateTime(serial, out _) is { } rangeError)
        {
            return ComputedValue.Error(rangeError);
        }

        return ComputedValue.Number(DateSerial.TimeOfDaySeconds(serial) / 3600);
    }
}

[MemoryPackable]
public sealed partial record Minute(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (DateSerial.ToDateTime(serial, out _) is { } rangeError)
        {
            return ComputedValue.Error(rangeError);
        }

        return ComputedValue.Number(DateSerial.TimeOfDaySeconds(serial) % 3600 / 60);
    }
}

[MemoryPackable]
public sealed partial record Second(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (DateSerial.ToDateTime(serial, out _) is { } rangeError)
        {
            return ComputedValue.Error(rangeError);
        }

        return ComputedValue.Number(DateSerial.TimeOfDaySeconds(serial) % 60);
    }
}
