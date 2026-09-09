using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Working-day functions: NETWORKDAYS(.INTL) count working days in a closed range; WORKDAY(.INTL) step a
// number of working days from a start. Holidays arrive as a range and are flattened with the shared
// ArgumentFlattening helper. The weekend pattern (which weekdays are non-working) is a WeekendSchedule,
// built either from the fixed Sat/Sun default, a weekend number (1..7, 11..17), or a 7-char "0000011" mask.
//
// Both walks run on SERIALS and derive each weekday from the serial through the central map, so Excel's
// phantom 1900-02-29 (serial 60) is a day of the walk like any other; below serial 61 the walk deliberately
// follows the REAL calendar rather than Aspose's answers — see WorkdayMath.IsWorkingSerial for the ruling and
// the evidence.

/// <summary>
/// Which days of the week are non-working. Indexed by <see cref="DayOfWeek"/> (Sunday = 0 .. Saturday = 6).
/// </summary>
internal readonly struct WeekendSchedule(bool[] weekend)
{
    /// <summary>The default Saturday+Sunday weekend (weekend number 1) used by NETWORKDAYS/WORKDAY.</summary>
    public static WeekendSchedule Default { get; } =
        new([true, false, false, false, false, false, true]); // Sunday and Saturday

    public bool IsWeekend(DayOfWeek dayOfWeek) => weekend[(int)dayOfWeek];

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

    /// <summary>
    /// Whether a whole-day serial is a working day: its weekday is not in the schedule's weekend and the
    /// serial is not a holiday. The weekday is the REAL (Gregorian) one the central map defines — serial 0 is
    /// 1899-12-31, a Sunday; serial 1 is 1900-01-01, a Monday; and Excel's phantom serial 60 repeats serial
    /// 59's 1900-02-28, a Wednesday, so both 59 and 60 count as working days.
    /// </summary>
    /// <remarks>
    /// <b>Do not "fix" this to Aspose's answers below serial 61.</b> By CONTROLLER RULING under the user's
    /// exception (2026-09-09) the working-day family walks the REAL calendar for every serial, and only
    /// serials ≥ 61 are pinned to the oracle. Aspose.Cells 26.6.0 has no derivable rule in the
    /// January–February 1900 window — the phase verifier fitted fourteen candidate rules and the best still
    /// missed 17 of 580 rows — and its answers there contradict one another, so there is nothing consistent to
    /// reproduce (every value below MEASURED on Aspose.Cells 26.6.0, 2026-09-09, PLAIN cell entry):
    /// <list type="number">
    /// <item><c>WORKDAY(6,4)</c> = <c>WORKDAY(6,5)</c> = 12 — the 4th and the 5th working day from the same
    /// start are the same day, so the result is not a function of <c>days</c>.</item>
    /// <item><c>WORKDAY.INTL(1,5,"1000000")</c> = <c>WORKDAY.INTL(1,6,"1000000")</c> = 7 — the same collapse
    /// under a one-day weekend.</item>
    /// <item><c>WORKDAY(58,4)</c> = 64, and serial 64 is a weekend on BOTH calendars
    /// (<c>WEEKDAY(64)</c> = 1, <c>NETWORKDAYS(64,64)</c> = 0, <c>TEXT(64,"dddd")</c> = Sunday): the answer is
    /// a day Aspose itself does not count as a working day.</item>
    /// <item><c>NETWORKDAYS</c> is not additive over a split range: <c>NETWORKDAYS(58,62)</c> = 4 while
    /// <c>NETWORKDAYS(58,58)</c> + <c>NETWORKDAYS(59,61)</c> + <c>NETWORKDAYS(62,62)</c> = 1 + 3 + 1 = 5, so
    /// no per-serial working-day indicator — this one included — can produce that count.</item>
    /// </list>
    /// Serials ≥ 61 are untouched by the exception: from 1900-03-01 on the map is the identity, and this walk
    /// reproduces the oracle exactly there (see the <c>Modern_*</c> pins in <c>DateEpochTests</c>).
    /// </remarks>
    public static bool IsWorkingSerial(
        int serial,
        in WeekendSchedule schedule,
        HashSet<int> holidays
    ) =>
        !schedule.IsWeekend(DateSerial.ToDateTimeUnchecked(serial).DayOfWeek)
        && !holidays.Contains(serial);

    /// <summary>
    /// Inclusive count of working days between the two serials; negative when end precedes start. The callers
    /// have already range-checked both ends through <see cref="DateSerial.ToDateTime"/>.
    /// </summary>
    public static int CountNetworkDays(
        double startSerial,
        double endSerial,
        in WeekendSchedule schedule,
        HashSet<int> holidays
    )
    {
        // The SERIAL is the source of truth and the calendar day is derived from it, never the reverse: under
        // Excel's epoch serials 59 and 60 both denote 1900-02-28 and no DateTime yields 60, so a counter
        // walked alongside a DateTime desynchronizes across that boundary and the holiday lookups silently
        // read the wrong day. Walking serials also keeps the phantom day in the count, which is what makes
        // NETWORKDAYS(59,61) = 3 instead of 2.
        var start = (int)Math.Floor(startSerial);
        var end = (int)Math.Floor(endSerial);
        var sign = 1;

        if (start > end)
        {
            (start, end) = (end, start);
            sign = -1;
        }

        var count = 0;

        for (var serial = start; serial <= end; serial++)
        {
            if (IsWorkingSerial(serial, schedule, holidays))
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

        // The map is asked only to police the range (negative or past 9999-12-31 → #NUM!); the walk itself
        // runs on the serials, per WorkdayMath.CountNetworkDays.
        if (DateSerial.ToDateTime(startSerial, out _) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out _) is { } endRange)
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
            WorkdayMath.CountNetworkDays(startSerial, endSerial, WeekendSchedule.Default, holidays)
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

        // The map only polices the range; the count walks the serials (WorkdayMath.CountNetworkDays).
        if (DateSerial.ToDateTime(startSerial, out _) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out _) is { } endRange)
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
            WorkdayMath.CountNetworkDays(startSerial, endSerial, schedule, holidays)
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

        // The map only polices the range; the step walks the serials (see Advance).
        if (DateSerial.ToDateTime(startSerial, out _) is { } startRange)
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

        return Advance(startSerial, daysArg, WeekendSchedule.Default, holidays);
    }

    /// <summary>
    /// The serial <paramref name="daysArg"/> working days from <paramref name="startSerial"/> (the start
    /// itself is never counted); a negative count walks backward. Stepping below serial 0 or past
    /// <see cref="DateSerial.MaxSerial"/> is <see cref="Error.Num"/>.
    /// </summary>
    /// <remarks>
    /// Walks SERIALS and derives each weekday from the serial — see
    /// <see cref="WorkdayMath.IsWorkingSerial"/> for why, and for the controller ruling that keeps this walk on
    /// the real calendar below serial 61 instead of chasing Aspose's self-contradicting composite.
    /// </remarks>
    internal static ComputedValue Advance(
        double startSerial,
        double daysArg,
        in WeekendSchedule schedule,
        HashSet<int> holidays
    )
    {
        var serial = (int)Math.Floor(startSerial);
        var days = (int)Math.Truncate(daysArg);

        if (days == 0)
        {
            // Zero days never moves, so it answers the start even when every day is a weekend. Measured on
            // Aspose.Cells 26.6.0 (2026-09-09, PLAIN): WORKDAY.INTL(45366,0,"1111111") = 45366 and
            // WORKDAY.INTL(1,0,"1111111") = 1 — the all-weekend guard below must NOT run first.
            return ComputedValue.Number(serial);
        }

        if (schedule.AllWeekend)
        {
            // No day to land on. Measured on Aspose.Cells 26.6.0 (2026-09-09, PLAIN):
            // WORKDAY.INTL(45366,5,"1111111") = WORKDAY.INTL(1,5,"1111111") =
            // WORKDAY.INTL(DATE(2012,1,1),30,"1111111") = #VALUE!, not the #NUM! the Microsoft page implies.
            return ComputedValue.Error(Error.Value);
        }

        var step = days > 0 ? 1 : -1;
        var remaining = Math.Abs(days);

        while (remaining > 0)
        {
            serial += step;

            if (serial < 0 || serial > DateSerial.MaxSerial)
            {
                return ComputedValue.Error(Error.Num);
            }

            if (WorkdayMath.IsWorkingSerial(serial, schedule, holidays))
            {
                remaining--;
            }
        }

        return ComputedValue.Number(serial);
    }
}

[MemoryPackable]
public sealed partial record WorkdayIntl(Expression[] Arguments) : Function
{
    // WORKDAY.INTL(start, days, [weekend], [holidays]) — WORKDAY with a configurable weekend. An invalid
    // weekend number (e.g. 0) → #NUM!; an all-weekend schedule has no day to land on → #VALUE! (measured).
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

        // The map only polices the range; the step walks the serials (see Workday.Advance).
        if (DateSerial.ToDateTime(startSerial, out _) is { } startRange)
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

        return Workday.Advance(startSerial, daysArg, schedule, holidays);
    }
}
