using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Working-day functions: NETWORKDAYS(.INTL) count working days in a closed range; WORKDAY(.INTL) step a
// number of working days from a start. Holidays arrive as a range and are flattened with the shared
// ArgumentFlattening helper. The weekend pattern (which weekdays are non-working) is a WeekendSchedule,
// built either from the fixed Sat/Sun default, a weekend number (1..7, 11..17), or a 7-char "0000011" mask.

/// <summary>
/// Which days of the week are non-working. Indexed by <see cref="DayOfWeek"/> (Sunday = 0 .. Saturday = 6).
/// </summary>
internal readonly struct WeekendSchedule(bool[] weekend)
{
    /// <summary>The default Saturday+Sunday weekend (weekend number 1) used by NETWORKDAYS/WORKDAY.</summary>
    public static WeekendSchedule Default { get; } =
        new([true, false, false, false, false, false, true]); // Sunday and Saturday

    public bool IsWeekend(DateTime date) => weekend[(int)date.DayOfWeek];

    public bool AllWeekend => Array.TrueForAll(weekend, day => day);

    /// <summary>
    /// Builds a schedule from the optional <c>[weekend]</c> argument. A 7-character text is the mask
    /// (Monday→Sunday, '1' = non-working; malformed → <see cref="Error.Value"/>); otherwise a weekend
    /// number 1..7 (two-day weekends) or 11..17 (single-day), anything else → <see cref="Error.Num"/>.
    /// </summary>
    public static Error? FromArgument(in ComputedValue argument, out WeekendSchedule schedule)
    {
        schedule = default;

        if (argument.TryGetText(out var text))
        {
            if (text.Length != 7)
            {
                return Error.Value;
            }

            var mask = new bool[7];

            for (var i = 0; i < 7; i++)
            {
                switch (text[i])
                {
                    case '1':
                        mask[(i + 1) % 7] = true; // position 0 = Monday (DayOfWeek 1) … position 6 = Sunday (0)
                        break;
                    case '0':
                        break;
                    default:
                        return Error.Value;
                }
            }

            schedule = new WeekendSchedule(mask);
            return null;
        }

        if (argument.CoerceToNumber(out var number) is { } error)
        {
            return error;
        }

        var weekendNumber = (int)Math.Truncate(number);
        var days = new bool[7];

        switch (weekendNumber)
        {
            case >= 1 and <= 7:
                var first = (5 + weekendNumber) % 7; // 1 → Saturday(6)/Sunday(0), shifting forward
                days[first] = true;
                days[(first + 1) % 7] = true;
                break;

            case >= 11
            and <= 17:
                days[weekendNumber - 11] = true; // 11 → Sunday(0) … 17 → Saturday(6)
                break;

            default:
                return Error.Num;
        }

        schedule = new WeekendSchedule(days);
        return null;
    }
}

internal static class WorkdayMath
{
    /// <summary>Collects the optional holidays argument (a range/list) into a set of whole-day serials.</summary>
    public static Error? CollectHolidays(
        Expression holidayArgument,
        EvaluationContext context,
        HashSet<int> holidays
    )
    {
        foreach (var value in ArgumentFlattening.ExpandComputedValues(holidayArgument, context))
        {
            if (value.Kind == ComputedValueKind.Blank)
            {
                continue;
            }

            if (value.CoerceToNumber(out var serial) is { } error)
            {
                return error;
            }

            if (serial < 0d)
            {
                return Error.Num;
            }

            holidays.Add((int)Math.Floor(serial));
        }

        return null;
    }

    /// <summary>Inclusive count of working days between the two dates; negative when end precedes start.</summary>
    public static int CountNetworkDays(
        DateTime start,
        DateTime end,
        in WeekendSchedule schedule,
        HashSet<int> holidays
    )
    {
        var sign = 1;

        if (start > end)
        {
            (start, end) = (end, start);
            sign = -1;
        }

        var count = 0;

        // Walk the serial alongside the DateTime instead of calling ToOADate() every step: whole-day
        // DateTime values advance the OADate by exactly 1 per AddDays(1), so an int++ is bit-exact and
        // skips the per-day double conversion. DayOfWeek still needs the DateTime (WeekendSchedule).
        var serial = (int)start.ToOADate();

        for (var day = start; day <= end; day = day.AddDays(1), serial++)
        {
            if (!schedule.IsWeekend(day) && !holidays.Contains(serial))
            {
                count++;
            }
        }

        return count * sign;
    }
}

[MemoryPackable]
public sealed partial record NetworkDays(Expression[] Arguments) : Function
{
    // NETWORKDAYS(start, end, [holidays]) — inclusive of both ends; Saturdays/Sundays and holidays excluded.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet reference (start/end cell or the holidays range) is a structural #REF! — a ghost
        // holidays range would otherwise be read as empty and silently ignored.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var endSerial) is { } endError)
        {
            return ComputedValue.Error(endError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out var end) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        var holidays = new HashSet<int>();

        if (
            Arguments.Length == 3
            && WorkdayMath.CollectHolidays(Arguments[2], context, holidays) is { } holidayError
        )
        {
            return ComputedValue.Error(holidayError);
        }

        return ComputedValue.Number(
            WorkdayMath.CountNetworkDays(start.Date, end.Date, WeekendSchedule.Default, holidays)
        );
    }
}

[MemoryPackable]
public sealed partial record NetworkDaysIntl(Expression[] Arguments) : Function
{
    // NETWORKDAYS.INTL(start, end, [weekend], [holidays]) — like NETWORKDAYS with a configurable weekend.
    // An all-weekend schedule ("1111111") legitimately yields 0.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet reference (start/end cell or the holidays range) is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var endSerial) is { } endError)
        {
            return ComputedValue.Error(endError);
        }

        var schedule = WeekendSchedule.Default;

        if (
            Arguments.Length >= 3
            && WeekendSchedule.FromArgument(Arguments[2].Evaluate(context), out schedule)
                is { } weekendError
        )
        {
            return ComputedValue.Error(weekendError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out var end) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        var holidays = new HashSet<int>();

        if (
            Arguments.Length == 4
            && WorkdayMath.CollectHolidays(Arguments[3], context, holidays) is { } holidayError
        )
        {
            return ComputedValue.Error(holidayError);
        }

        return ComputedValue.Number(
            WorkdayMath.CountNetworkDays(start.Date, end.Date, schedule, holidays)
        );
    }
}

[MemoryPackable]
public sealed partial record Workday(Expression[] Arguments) : Function
{
    // WORKDAY(start, days, [holidays]) — the date `days` working days from start (the start itself is not
    // counted); negative days walk backward.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet reference (start cell or the holidays range) is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var daysArg) is { } daysError)
        {
            return ComputedValue.Error(daysError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        var holidays = new HashSet<int>();

        if (
            Arguments.Length == 3
            && WorkdayMath.CollectHolidays(Arguments[2], context, holidays) is { } holidayError
        )
        {
            return ComputedValue.Error(holidayError);
        }

        return Advance(start.Date, daysArg, WeekendSchedule.Default, holidays);
    }

    internal static ComputedValue Advance(
        DateTime start,
        double daysArg,
        in WeekendSchedule schedule,
        HashSet<int> holidays
    )
    {
        if (schedule.AllWeekend)
        {
            return ComputedValue.Error(Error.Num);
        }

        var days = (int)Math.Truncate(daysArg);

        if (days == 0)
        {
            return ComputedValue.Number(start.ToOADate());
        }

        var step = days > 0 ? 1 : -1;
        var remaining = Math.Abs(days);
        var current = start;
        // See CountNetworkDays: walk the serial alongside current instead of calling ToOADate() per step.
        var serial = (int)start.ToOADate();

        try
        {
            while (remaining > 0)
            {
                current = current.AddDays(step);
                serial += step;

                if (!schedule.IsWeekend(current) && !holidays.Contains(serial))
                {
                    remaining--;
                }
            }
        }
        catch (ArgumentOutOfRangeException)
        {
            return ComputedValue.Error(Error.Num);
        }

        return ComputedValue.Number(current.ToOADate());
    }
}

[MemoryPackable]
public sealed partial record WorkdayIntl(Expression[] Arguments) : Function
{
    // WORKDAY.INTL(start, days, [weekend], [holidays]) — WORKDAY with a configurable weekend. An invalid
    // weekend number (e.g. 0) → #NUM!; an all-weekend schedule has no day to land on → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet reference (start cell or the holidays range) is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var daysArg) is { } daysError)
        {
            return ComputedValue.Error(daysError);
        }

        var schedule = WeekendSchedule.Default;

        if (
            Arguments.Length >= 3
            && WeekendSchedule.FromArgument(Arguments[2].Evaluate(context), out schedule)
                is { } weekendError
        )
        {
            return ComputedValue.Error(weekendError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        var holidays = new HashSet<int>();

        if (
            Arguments.Length == 4
            && WorkdayMath.CollectHolidays(Arguments[3], context, holidays) is { } holidayError
        )
        {
            return ComputedValue.Error(holidayError);
        }

        return Workday.Advance(start.Date, daysArg, schedule, holidays);
    }
}
