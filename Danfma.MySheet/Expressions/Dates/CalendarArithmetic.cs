using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Calendar arithmetic over serials: DAYS, DAYS360, EDATE, EOMONTH, WEEKDAY, WEEKNUM, ISOWEEKNUM, DATEDIF,
// YEARFRAC. Serials are validated non-negative (negative → #NUM!); the day-count bases of YEARFRAC live in
// the shared DayCount helper (reused by the wave-6 bond functions).

[MemoryPackable]
public sealed partial record Days(Expression[] Arguments) : Function
{
    // DAYS(end_date, start_date) = whole days between them; the result may be negative (end before start).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var endSerial) is { } endError)
        {
            return ComputedValue.Error(endError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (DateSerial.ToDateTime(endSerial, out var endDate) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        if (DateSerial.ToDateTime(startSerial, out var startDate) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        return ComputedValue.Number((endDate.Date - startDate.Date).Days);
    }
}

[MemoryPackable]
public sealed partial record Days360(Expression[] Arguments) : Function
{
    // DAYS360(start, end, [method]) — 30/360 day count. method FALSE/omitted = US (NASD), TRUE = European.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var endSerial) is { } endError)
        {
            return ComputedValue.Error(endError);
        }

        var european = false;

        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToBool(out european) is { } methodError
        )
        {
            return ComputedValue.Error(methodError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out var end) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        var days = european
            ? DayCount.Euro360Days(start.Date, end.Date)
            : UsDays360(start.Date, end.Date);

        return ComputedValue.Number(days);
    }

    // US (NASD) DAYS360 (support.microsoft.com + MS-OI29500 §18.17.7.79). Distinct from YEARFRAC basis 0:
    // when the end date is the last day of a month and the (adjusted) start day is < 30, the end rolls to
    // the 1st of the next month.
    private static int UsDays360(DateTime start, DateTime end)
    {
        int d1 = start.Day,
            m1 = start.Month,
            y1 = start.Year;
        int d2 = end.Day,
            m2 = end.Month,
            y2 = end.Year;

        if (d1 == DateTime.DaysInMonth(y1, m1))
        {
            d1 = 30;
        }

        if (d2 == DateTime.DaysInMonth(y2, m2))
        {
            if (d1 < 30)
            {
                d2 = 1;
                m2++;

                if (m2 > 12)
                {
                    m2 = 1;
                    y2++;
                }
            }
            else
            {
                d2 = 30;
            }
        }

        return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
    }
}

[MemoryPackable]
public sealed partial record EDate(Expression[] Arguments) : Function
{
    // EDATE(start, months) — same day-of-month `months` away, clamped to the month end (Jan 31 + 1 → Feb 28).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var monthsArg) is { } monthsError)
        {
            return ComputedValue.Error(monthsError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        return Shift(start.Date, monthsArg, out var shifted)
            ? ComputedValue.Number(DateSerial.FromDateTime(shifted))
            : ComputedValue.Error(Error.Num);
    }

    internal static bool Shift(DateTime date, double monthsArg, out DateTime shifted)
    {
        var months = Math.Truncate(monthsArg);

        if (double.IsNaN(months) || Math.Abs(months) > 120000d)
        {
            shifted = default;
            return false;
        }

        try
        {
            shifted = date.AddMonths((int)months);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            shifted = default;
            return false;
        }
    }
}

[MemoryPackable]
public sealed partial record EoMonth(Expression[] Arguments) : Function
{
    // EOMONTH(start, months) — the last day of the month `months` away from start.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var monthsArg) is { } monthsError)
        {
            return ComputedValue.Error(monthsError);
        }

        if (DateSerial.ToDateTime(startSerial, out var start) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (!EDate.Shift(start.Date, monthsArg, out var shifted))
        {
            return ComputedValue.Error(Error.Num);
        }

        var lastDay = new DateTime(
            shifted.Year,
            shifted.Month,
            DateTime.DaysInMonth(shifted.Year, shifted.Month)
        );

        return ComputedValue.Number(DateSerial.FromDateTime(lastDay));
    }
}

[MemoryPackable]
public sealed partial record Weekday(Expression[] Arguments) : Function
{
    // WEEKDAY(serial, [return_type]) — return_type 1/2/3 and 11..17 (see the reference table). Unknown
    // return_type → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var returnType = 1d;

        if (
            Arguments.Length == 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out returnType) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        if (DateSerial.ToDateTime(serial, out var date) is { } rangeError)
        {
            return ComputedValue.Error(rangeError);
        }

        var dow = (int)date.DayOfWeek; // Sunday = 0 .. Saturday = 6
        var type = (int)Math.Truncate(returnType);

        return type switch
        {
            1 or 17 => ComputedValue.Number(dow + 1), // Sunday = 1 .. Saturday = 7
            2 or 11 => ComputedValue.Number((dow + 6) % 7 + 1), // Monday = 1 .. Sunday = 7
            3 => ComputedValue.Number((dow + 6) % 7), // Monday = 0 .. Sunday = 6
            >= 12 and <= 16 => ComputedValue.Number((dow - (type - 10) % 7 + 7) % 7 + 1),
            _ => ComputedValue.Error(Error.Num),
        };
    }
}

[MemoryPackable]
public sealed partial record WeekNum(Expression[] Arguments) : Function
{
    // WEEKNUM(serial, [return_type]) — System 1 (week containing Jan 1 is week 1) for return_type
    // 1/2/11..17, and System 2 = ISO 8601 for return_type 21. Unknown return_type → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var returnType = 1d;

        if (
            Arguments.Length == 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out returnType) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        if (DateSerial.ToDateTime(serial, out var date) is { } rangeError)
        {
            return ComputedValue.Error(rangeError);
        }

        var type = (int)Math.Truncate(returnType);

        if (type == 21)
        {
            return ComputedValue.Number(ISOWeek.GetWeekOfYear(date));
        }

        // Week-start day as a .NET DayOfWeek index (Sunday = 0 .. Saturday = 6).
        var startDow = type switch
        {
            1 or 17 => 0, // Sunday
            2 or 11 => 1, // Monday
            12 => 2,
            13 => 3,
            14 => 4,
            15 => 5,
            16 => 6,
            _ => -1,
        };

        if (startDow < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var jan1 = new DateTime(date.Year, 1, 1);
        var offset = ((int)jan1.DayOfWeek - startDow + 7) % 7;
        var weekNum = (date.DayOfYear - 1 + offset) / 7 + 1;

        return ComputedValue.Number(weekNum);
    }
}

[MemoryPackable]
public sealed partial record IsoWeekNum(Expression[] Arguments) : Function
{
    // ISOWEEKNUM(serial) — ISO 8601 week number (weeks start Monday; week 1 contains the year's first
    // Thursday / January 4).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return DateSerial.ToDateTime(serial, out var date) is { } rangeError
            ? ComputedValue.Error(rangeError)
            : ComputedValue.Number(ISOWeek.GetWeekOfYear(date));
    }
}

[MemoryPackable]
public sealed partial record DateDif(Expression[] Arguments) : Function
{
    // DATEDIF(start, end, unit) — Lotus-compatibility function with the 6 documented units. start > end is
    // #NUM!. The "MD" unit is officially warned to be unreliable (may be negative/zero/inaccurate); its
    // classic Excel behavior is reproduced, not corrected.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var endSerial) is { } endError)
        {
            return ComputedValue.Error(endError);
        }

        if (Arguments[2].Evaluate(context).CoerceToText(out var unit) is { } unitError)
        {
            return ComputedValue.Error(unitError);
        }

        if (DateSerial.ToDateTime(startSerial, out var startDt) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out var endDt) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        var start = startDt.Date;
        var end = endDt.Date;

        if (start > end)
        {
            return ComputedValue.Error(Error.Num);
        }

        return unit.ToUpperInvariant() switch
        {
            "Y" => ComputedValue.Number(CompleteYears(start, end)),
            "M" => ComputedValue.Number(CompleteMonths(start, end)),
            "D" => ComputedValue.Number((end - start).Days),
            "MD" => ComputedValue.Number(DayDifferenceIgnoringMonths(start, end)),
            "YM" => ComputedValue.Number(MonthDifferenceIgnoringYears(start, end)),
            "YD" => ComputedValue.Number(DayDifferenceIgnoringYears(start, end)),
            _ => ComputedValue.Error(Error.Num),
        };
    }

    private static int CompleteYears(DateTime start, DateTime end)
    {
        var years = end.Year - start.Year;

        if (end.Month < start.Month || (end.Month == start.Month && end.Day < start.Day))
        {
            years--;
        }

        return years;
    }

    private static int CompleteMonths(DateTime start, DateTime end)
    {
        var months = (end.Year - start.Year) * 12 + (end.Month - start.Month);

        if (end.Day < start.Day)
        {
            months--;
        }

        return months;
    }

    private static int DayDifferenceIgnoringMonths(DateTime start, DateTime end)
    {
        if (end.Day >= start.Day)
        {
            return end.Day - start.Day;
        }

        var previousMonth = end.AddMonths(-1);
        var daysInPreviousMonth = DateTime.DaysInMonth(previousMonth.Year, previousMonth.Month);

        return daysInPreviousMonth - start.Day + end.Day;
    }

    private static int MonthDifferenceIgnoringYears(DateTime start, DateTime end)
    {
        var months = end.Month - start.Month;

        if (end.Day < start.Day)
        {
            months--;
        }

        return months < 0 ? months + 12 : months;
    }

    private static int DayDifferenceIgnoringYears(DateTime start, DateTime end)
    {
        var anchoredStart = start.AddYears(end.Year - start.Year);

        if (anchoredStart > end)
        {
            anchoredStart = anchoredStart.AddYears(-1);
        }

        return (end - anchoredStart).Days;
    }
}

[MemoryPackable]
public sealed partial record YearFrac(Expression[] Arguments) : Function
{
    // YEARFRAC(start, end, [basis]) — year fraction on one of the 5 day-count bases (0..4). Arguments are
    // truncated to whole-day serials (MS), the pair is ordered (result is non-negative), and the day-count
    // machinery is shared with wave 6 via DayCount.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var startSerial) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var endSerial) is { } endError)
        {
            return ComputedValue.Error(endError);
        }

        var basisArg = 0d;

        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out basisArg) is { } basisError
        )
        {
            return ComputedValue.Error(basisError);
        }

        var basis = (int)Math.Truncate(basisArg);

        if (basis is < 0 or > 4)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (DateSerial.ToDateTime(startSerial, out var startDt) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out var endDt) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        var start = startDt.Date;
        var end = endDt.Date;

        if (start == end)
        {
            return ComputedValue.Number(0d);
        }

        if (start > end)
        {
            (start, end) = (end, start);
        }

        return ComputedValue.Number(DayCount.YearFraction(start, end, basis));
    }
}
