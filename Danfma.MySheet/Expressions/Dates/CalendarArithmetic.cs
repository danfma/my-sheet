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

        // The map supplies the range policy only. The COUNT is a serial subtraction: the two serials 59 and
        // 60 both map to 1900-02-28, so a DateTime difference silently swallows the phantom day whenever the
        // span straddles it (measured DAYS(60,59) = 1, DAYS(61,59) = 2, DAYS(59,61) = -2).
        if (DateSerial.ToDateTime(endSerial, out _) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        if (DateSerial.ToDateTime(startSerial, out _) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        return ComputedValue.Number(Math.Truncate(endSerial) - Math.Truncate(startSerial));
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

        // The 30/360 arithmetic reads the LOTUS calendar: serial 60 is 1900-02-29 and February 1900 has 29
        // days, so serial 59 is NOT that month's last day (measured DAYS360(59,61) = 3).
        if (
            DateSerial.TryGetCalendar(
                startSerial,
                phantomFeb29: true,
                out var y1,
                out var m1,
                out var d1
            ) is
            { } startRange
        )
        {
            return ComputedValue.Error(startRange);
        }

        if (
            DateSerial.TryGetCalendar(
                endSerial,
                phantomFeb29: true,
                out var y2,
                out var m2,
                out var d2
            ) is
            { } endRange
        )
        {
            return ComputedValue.Error(endRange);
        }

        // MEASURED and not derivable from the rules below: the phantom day counts as day 30 when it OPENS the
        // span and as day 29 when it closes it (DAYS360(60,61) = DAYS360(60,61,TRUE) = 1 against
        // DAYS360(59,60) = DAYS360(59,60,TRUE) = 1), and a span from the phantom day to itself is 0 rather
        // than the -1 those two roles would otherwise produce (DAYS360(60,60) = DAYS360(60.5,60.9) = 0).
        if (Math.Truncate(startSerial) == Math.Truncate(endSerial))
        {
            return ComputedValue.Number(0d);
        }

        if (DateSerial.IsPhantomFeb29(startSerial))
        {
            d1 = 30;
        }

        var days = european
            ? DayCount.Euro360Days(y1, m1, d1, y2, m2, d2)
            : UsDays360(y1, m1, d1, y2, m2, d2);

        return ComputedValue.Number(days);
    }

    // US (NASD) DAYS360 as MEASURED on Aspose.Cells 26.6.0 (PLAIN entry). Two things separate it from
    // YEARFRAC's basis 0 in DayCount.Nasd360Days:
    //   * the February pull runs BEFORE the end-31 test, so a February-end start DOES drag a day-31 end down
    //     with it — DAYS360(DATE(2023,2,28),DATE(2023,3,31)) = 30 against YEARFRAC(…,0) = 31/360; and
    //   * the month length is the LOTUS one, so 1900-02-29 is February's last day and 1900-02-28 is not.
    // There is NO "end rolls to the 1st of the next month" rule, which support.microsoft.com and
    // MS-OI29500 §18.17.7.79 both describe and Aspose does not have: measured
    // DAYS360(DATE(2011,1,1),DATE(2011,4,30)) = 119 (not 120), (DATE(2024,1,16),DATE(2024,2,29)) = 43
    // (not 45), (DATE(2011,1,15),DATE(2011,9,30)) = 255.
    private static int UsDays360(int y1, int m1, int d1, int y2, int m2, int d2)
    {
        if (d1 == 31)
        {
            d1 = 30;
        }

        if (m1 == 2 && d1 == DateSerial.LotusDaysInMonth(y1, 2))
        {
            d1 = 30;
        }

        if (d2 == 31 && d1 >= 30)
        {
            d2 = 30;
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

        if (!Shift(start.Date, monthsArg, out var shifted))
        {
            return ComputedValue.Error(Error.Num);
        }

        var serial = DateSerial.FromDateTime(shifted);

        // A shift that lands before the day zero has no serial at all: #NUM! (measured EDATE(0,-1),
        // EDATE(1,-1), EDATE(2,-1), EDATE(60,-2)), where the OA epoch used to hand back a negative serial
        // that every other date function then rejected anyway. EDATE(31,-1) = 0 stays valid — the day zero
        // itself is a serial.
        return serial < 0d ? ComputedValue.Error(Error.Num) : ComputedValue.Number(serial);
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
    // EOMONTH(start, months) — the last day of the month `months` away from start. Only the start's MONTH
    // matters, so it is read through TryGetCalendar and the shift runs from the 1st: that is what puts the day
    // zero in January 1900 (measured EOMONTH(0,0) = EOMONTH(0.5,0) = 31, not the 0 that 1899-12-31 would give)
    // and it also spares the shift the day-of-month clamping EDATE needs.
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

        if (
            DateSerial.TryGetCalendar(
                startSerial,
                phantomFeb29: false,
                out var year,
                out var month,
                out _
            ) is
            { } startRange
        )
        {
            return ComputedValue.Error(startRange);
        }

        // MEASURED and NOT derivable from the month arithmetic: a start inside the day-zero window [0, 1)
        // with a negative (truncated) month count is #NUM! — EOMONTH(0,-1) and EOMONTH(0.5,-1) — even though
        // the identical shift from serial 1 is perfectly valid, EOMONTH(1,-1) = 0, and even though a count
        // that truncates to zero is fine, EOMONTH(0,-0.5) = 31. Aspose.Cells 26.6.0 has no single rule here,
        // so this is a pinned special case rather than a consequence of anything.
        if (startSerial < 1d && Math.Truncate(monthsArg) < 0d)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (!EDate.Shift(new DateTime(year, month, 1), monthsArg, out var shifted))
        {
            return ComputedValue.Error(Error.Num);
        }

        var lastDay = new DateTime(
            shifted.Year,
            shifted.Month,
            DateTime.DaysInMonth(shifted.Year, shifted.Month)
        );

        var serial = DateSerial.FromDateTime(lastDay);

        // Before the day zero there is no serial (measured EOMONTH(31,-2) = #NUM!, while EOMONTH(31,-1) = 0).
        return serial < 0d ? ComputedValue.Error(Error.Num) : ComputedValue.Number(serial);
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

        // The range policy still comes from the map, but the weekday itself does NOT: the Lotus mod-7 walk
        // skips the phantom day, so serial 60 repeats serial 59's Tuesday, and the shifted map's own
        // DayOfWeek would be a day off for every serial in the 1900 window.
        if (DateSerial.ToDateTime(serial, out _) is { } rangeError)
        {
            return ComputedValue.Error(rangeError);
        }

        var dow = (int)DateSerial.LotusDayOfWeek(serial); // Sunday = 0 .. Saturday = 6
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
    // #NUM!. "Y", "M" and "YM" are whole calendar counts read off the DateTime map; "D", "MD" and "YD" count
    // SERIALS from an anchor (see MonthAnchor/YearAnchor), because a Gregorian difference is a day short over
    // the phantom 1900-02-29. Measured on Aspose.Cells 26.6.0, PLAIN entry.
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
        var startDay = Math.Truncate(startSerial);
        var endDay = Math.Truncate(endSerial);

        // The pair is ordered on the SERIALS, not on the mapped dates: 60 and 59 both map to 1900-02-28, and
        // measured DATEDIF(60,59,"d") and DATEDIF(61,60,"d") are #NUM! while DATEDIF(0,0,"d") is 0.
        if (startDay > endDay)
        {
            return ComputedValue.Error(Error.Num);
        }

        // The three calendar units read the map; the three day units subtract serials from an ANCHOR — the
        // start pushed forward by the whole months (or years) the span contains, which is the date the
        // leftover days are counted from. Anchoring is what makes the day units agree with Excel across the
        // phantom day (measured DATEDIF(60,61,"md") = 2, DATEDIF(31,60,"md") = 29) and away from it
        // (DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"md") = 1, where a plain "borrow the previous month's
        // length" gives -1).
        return unit.ToUpperInvariant() switch
        {
            "Y" => ComputedValue.Number(CompleteYears(start, end)),
            "M" => ComputedValue.Number(CompleteMonths(start, end)),
            "D" => ComputedValue.Number(endDay - startDay),
            "MD" => ComputedValue.Number(endDay - MonthAnchor(start, end)),
            "YM" => ComputedValue.Number(MonthDifferenceIgnoringYears(start, end)),
            "YD" => ComputedValue.Number(endDay - YearAnchor(start, end)),
            _ => ComputedValue.Error(Error.Num),
        };
    }

    /// <summary>
    /// The serial of the start pushed forward by every WHOLE month the span contains — the anchor "MD"
    /// counts its leftover days from. <see cref="DateTime.AddMonths"/> clamps to the target month's last day,
    /// which is what makes a day-31 start land on February 29 rather than rolling into March.
    /// </summary>
    private static double MonthAnchor(DateTime start, DateTime end) =>
        DateSerial.FromDateTime(start.AddMonths(CompleteMonths(start, end)));

    /// <summary>
    /// The serial of the start pushed forward by every WHOLE year the span contains — the anchor "YD" counts
    /// its leftover days from. <see cref="DateTime.AddYears"/> clamps February 29 to February 28.
    /// </summary>
    private static double YearAnchor(DateTime start, DateTime end) =>
        DateSerial.FromDateTime(start.AddYears(CompleteYears(start, end)));

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

    private static int MonthDifferenceIgnoringYears(DateTime start, DateTime end)
    {
        var months = end.Month - start.Month;

        if (end.Day < start.Day)
        {
            months--;
        }

        return months < 0 ? months + 12 : months;
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

        // The map supplies the range policy; the day counting itself stays on the serials, so that bases 1-3
        // span the phantom day and bases 0/4 can read it as February 29 (measured YEARFRAC(1,61,1) = 60/365,
        // (60,61,0) = 2/360).
        if (DateSerial.ToDateTime(startSerial, out _) is { } startRange)
        {
            return ComputedValue.Error(startRange);
        }

        if (DateSerial.ToDateTime(endSerial, out _) is { } endRange)
        {
            return ComputedValue.Error(endRange);
        }

        var start = Math.Floor(startSerial);
        var end = Math.Floor(endSerial);

        // Measured YEARFRAC(60,60,0) = YEARFRAC(59,59,0) = 0: the same whole day on both ends is no time at
        // all, ahead of the phantom day's asymmetric roles in the 30/360 count.
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
