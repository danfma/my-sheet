using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Component extraction from a date serial: YEAR/MONTH/DAY read the calendar part through
// DateSerial.TryGetCalendar, which is the DateTime map plus Excel's day-zero rule (serial 0 is 1900/1/0) and
// WITHOUT the phantom 1900-02-29 (serial 60 reads back as February 28, measured); HOUR/MINUTE/SECOND read the
// time-of-day fraction (rounded to the nearest second, matching Excel). A negative serial is out of range →
// #NUM!. Text-of-date arguments are NOT accepted (only numeric serials / numeric text that CoerceToNumber
// already parses) — see the date namespace note.

[MemoryPackable]
public sealed partial record Year(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (
            DateSerial.TryGetCalendar(serial, phantomFeb29: false, out var year, out _, out _) is
            { } rangeError
        )
        {
            return ComputedValue.Error(rangeError);
        }

        return ComputedValue.Number(year);
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

        if (
            DateSerial.TryGetCalendar(serial, phantomFeb29: false, out _, out var month, out _) is
            { } rangeError
        )
        {
            return ComputedValue.Error(rangeError);
        }

        return ComputedValue.Number(month);
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

        if (
            DateSerial.TryGetCalendar(serial, phantomFeb29: false, out _, out _, out var day) is
            { } rangeError
        )
        {
            return ComputedValue.Error(rangeError);
        }

        return ComputedValue.Number(day);
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
