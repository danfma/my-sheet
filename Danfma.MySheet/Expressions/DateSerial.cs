namespace Danfma.MySheet.Expressions;

/// <summary>
/// Excel serial-date arithmetic for the date/time functions. A date IS a <c>double</c> serial (fiel ao
/// Excel; no dedicated value kind): the integer part counts days and the fraction is the time of day.
/// Serial 1 is 1900-01-01 and serial 0 is Excel's "day zero" 1900-01-00, so the map is .NET's OLE
/// Automation date (<see cref="DateTime.FromOADate"/> / <see cref="DateTime.ToOADate"/>, epoch 1899-12-30)
/// shifted one day below 1900-03-01 and unshifted from 1900-03-01 on, where the two systems agree.
///
/// Excel is really THREE calendars over the January–February 1900 window, and each function is pinned to
/// the one it uses (all three MEASURED on Aspose.Cells 26.6.0, the oracle for "Excel"):
/// <list type="number">
/// <item><b>The DateTime map</b> — <see cref="ToDateTimeUnchecked"/> and its inverse
/// <see cref="FromDateTime"/>: 0 → 1899-12-31, 1 → 1900-01-01, 59 → 1900-02-28, 60 → 1900-02-28 (the
/// phantom day COLLAPSES onto Feb 28) and 61 → 1900-03-01. The inverse never yields 60, so no
/// <see cref="DateTime"/> denotes the phantom day. Used by YEAR/MONTH/DAY (through
/// <see cref="TryGetCalendar"/>, which adds the day-zero rule), EDATE, EOMONTH, WEEKNUM, ISOWEEKNUM,
/// NETWORKDAYS, the bond/coupon family, DATEDIF's calendar units, <c>DATE</c>'s inverse, the xlsx loader
/// and the volatile clock.</item>
/// <item><b>The Lotus weekday</b> — <see cref="LotusDayOfWeek"/>: <c>((⌊s⌋ − 1) mod 7) + 1</c>, with 60
/// collapsed onto 59. Used TODAY by WEEKDAY only; <c>TEXT</c>'s <c>ddd</c>/<c>dddd</c> is meant to join it
/// and still reads <see cref="DateTime.DayOfWeek"/> (Phase 9's TEXT item owns that move). WORKDAY and
/// NETWORKDAYS deliberately do NOT use it: by controller ruling under the user's exception they walk the
/// REAL calendar below serial 61, because Aspose has no consistent rule there. This is NOT the
/// <see cref="DateTime.DayOfWeek"/> of the mapped value, which the one-day shift moves.</item>
/// <item><b>The Lotus calendar</b> — serial 60 IS 1900-02-29 and February 1900 has 29 days, exposed by
/// <c>phantomFeb29: true</c>. Its intended consumers are the number formatter (<c>TEXT</c> and cell display)
/// and the 30/360 arithmetic of DAYS360/YEARFRAC; none of the three reads it yet — the pins for all of them
/// are still red and Phase 9's remaining items own the migration. Everything that COUNTS days counts
/// serials instead.</item>
/// </list>
///
/// Functions treat a negative serial as out of range → <c>#NUM!</c>.
/// </summary>
internal static class DateSerial
{
    /// <summary>Largest representable serial: 9999-12-31 23:59:59 (the OADate upper bound).</summary>
    public static readonly double MaxSerial = new DateTime(9999, 12, 31, 23, 59, 59).ToOADate();

    /// <summary>
    /// Excel's phantom 1900-02-29. No <see cref="DateTime"/> maps to it — <see cref="FromDateTime"/> never
    /// produces it and <see cref="ToDateTimeUnchecked"/> collapses it onto 1900-02-28 — so it is reachable
    /// only by arithmetic on serials and by <c>DATEVALUE</c>'s text parse.
    /// </summary>
    /// <remarks>
    /// Measured: the round trip <c>FromDateTime(ToDateTimeUnchecked(s))</c> is the identity for every serial
    /// EXCEPT the whole half-open day <c>[60, 61)</c>, which comes back as <c>s − 1</c> (60.5 → 59.5, not just
    /// the integer point). Anything that must preserve a serial across the phantom day has to carry the serial
    /// itself rather than a <see cref="DateTime"/> derived from it.
    /// </remarks>
    public const double PhantomFeb29Serial = 60d;

    /// <summary>
    /// The first date the Excel and OLE Automation epochs agree on. Below it Excel's serial is one less than
    /// the OADate, because Excel counts the phantom 1900-02-29 that the Gregorian calendar does not have.
    /// </summary>
    private static readonly DateTime FirstAlignedDate = new(1900, 3, 1);

    /// <summary>
    /// serial → <see cref="DateTime"/>. Returns <see cref="Error.Num"/> when the serial is negative or
    /// beyond the representable range; <c>null</c> on success.
    /// </summary>
    public static Error? ToDateTime(double serial, out DateTime dateTime)
    {
        if (double.IsNaN(serial) || serial < 0d || serial > MaxSerial + 1d)
        {
            dateTime = default;
            return Error.Num;
        }

        try
        {
            dateTime = ToDateTimeUnchecked(serial);
            return null;
        }
        catch (ArgumentException)
        {
            dateTime = default;
            return Error.Num;
        }
    }

    /// <summary>
    /// The serial → <see cref="DateTime"/> map itself, with no range policy of its own: the single place the
    /// epoch lives. <see cref="ToDateTime"/> wraps it with the negative/out-of-range →
    /// <see cref="Error.Num"/> guard; callers that answer a different error for an unrepresentable serial
    /// (<c>TEXT</c> answers <c>#VALUE!</c>) call this directly and keep their own policy. Throws
    /// <see cref="ArgumentException"/> outside the representable range, exactly as
    /// <see cref="DateTime.FromOADate"/> does.
    /// </summary>
    /// <remarks>
    /// THIS is the method the epoch lives in. Shifting <see cref="ToDateTime"/> instead compiles, passes its
    /// own tests, and silently leaves <c>TEXT</c> on the OLE-Automation epoch, because <c>TEXT</c> calls this
    /// method directly to keep its <c>#VALUE!</c> policy.
    /// </remarks>
    public static DateTime ToDateTimeUnchecked(double serial) =>
        // Below the phantom day Excel's serial trails the OADate by one, so +1 recovers the OADate; at 60 and
        // above the two agree, which is why 59 and 60 both land on 1900-02-28.
        DateTime.FromOADate(serial < PhantomFeb29Serial ? serial + 1d : serial);

    /// <summary>
    /// <see cref="DateTime"/> → serial, the inverse of <see cref="ToDateTimeUnchecked"/>. A date-only value
    /// yields an integer serial; <see cref="PhantomFeb29Serial"/> is never produced, and a date before
    /// 1899-12-31 yields a negative serial that the callers reject as <c>#NUM!</c>.
    /// </summary>
    public static double FromDateTime(DateTime dateTime)
    {
        var oaDate = dateTime.ToOADate();

        return dateTime < FirstAlignedDate ? oaDate - 1d : oaDate;
    }

    /// <summary>
    /// serial → calendar year/month/day, adding the two rules the raw <see cref="ToDateTime"/> map does not
    /// carry:
    /// <list type="bullet">
    /// <item>a serial in <c>[0, 1)</c> is Excel's day zero — 1900/1/<b>0</b>, not the 1899-12-31 the map
    /// returns (measured: <c>DAY(0)</c> = <c>DAY(0.999)</c> = 0, <c>MONTH(0)</c> = 1, <c>YEAR(0)</c> =
    /// 1900);</item>
    /// <item>a serial in <c>[60, 61)</c> is the phantom 1900/2/<b>29</b> for the callers that PRINT it or
    /// count 30/360 over it (<c>phantomFeb29: true</c>) and 1900/2/<b>28</b> for everyone else — YEAR, MONTH
    /// and DAY read the collapsed value (measured: <c>DAY(60)</c> = <c>DAY(60.5)</c> = 28).</item>
    /// </list>
    /// Returns <see cref="Error.Num"/> for an out-of-range serial, exactly as <see cref="ToDateTime"/> does.
    /// </summary>
    public static Error? TryGetCalendar(
        double serial,
        bool phantomFeb29,
        out int year,
        out int month,
        out int day
    )
    {
        year = 0;
        month = 0;
        day = 0;

        if (ToDateTime(serial, out var dateTime) is { } rangeError)
        {
            return rangeError;
        }

        if (serial < 1d)
        {
            (year, month, day) = (1900, 1, 0);
            return null;
        }

        if (phantomFeb29 && serial >= PhantomFeb29Serial && serial < PhantomFeb29Serial + 1d)
        {
            (year, month, day) = (1900, 2, 29);
            return null;
        }

        (year, month, day) = (dateTime.Year, dateTime.Month, dateTime.Day);
        return null;
    }

    /// <summary>
    /// The weekday Excel names for a serial: a plain mod-7 walk in which serial 1 (1900-01-01) is a
    /// <b>Sunday</b> — it was really a Monday, but the walk is anchored on the day-zero Saturday — and the
    /// phantom day is skipped, so serial 60 repeats serial 59's Tuesday. Deliberately NOT the
    /// <see cref="DateTime.DayOfWeek"/> of <see cref="ToDateTimeUnchecked"/>: the epoch shift moves that by a
    /// day, which would break every weekday answer in the 1900 window.
    /// </summary>
    public static DayOfWeek LotusDayOfWeek(double serial)
    {
        var day = (long)Math.Floor(serial);

        if (day == (long)PhantomFeb29Serial)
        {
            day--;
        }

        // The extra "+ 7" only guards a negative serial, which every caller rejects as #NUM! before asking.
        return (DayOfWeek)(int)(((day + 6L) % 7L + 7L) % 7L);
    }

    /// <summary>
    /// Builds a date serial from year/month/day with Excel's overflow rules (used by <c>DATE</c>, and the
    /// same normalization the other constructors rely on):
    /// <list type="bullet">
    /// <item>year 0..1899 → the value is added to 1900 (so <c>DATE(108,1,2)</c> is 2008-01-02);</item>
    /// <item>month outside 1..12 rolls the year (month 13 → January of the next year, month 0 → December of
    /// the previous year);</item>
    /// <item>day outside 1..(days in month) rolls the month (day 0 → last day of the previous month).</item>
    /// </list>
    /// year &lt; 0 or ≥ 10000, or any result outside 1900..9999, returns <see cref="Error.Num"/>. The floor is
    /// the day zero: <c>DATE(1900,1,0)</c> is serial 0 and <c>DATE(1900,1,-1)</c> would be −1, which the
    /// <c>result &lt; 0</c> guard below turns into <c>#NUM!</c>. Because this builds a proleptic-Gregorian
    /// date, <c>DATE(1900,2,29)</c> rolls to 1900-03-01 = serial <b>61</b> and never yields
    /// <see cref="PhantomFeb29Serial"/> — measured, and true of Excel too.
    /// </summary>
    public static Error? FromComponents(
        double yearArg,
        double monthArg,
        double dayArg,
        out double serial
    )
    {
        serial = 0d;

        var year = Math.Truncate(yearArg);

        if (double.IsNaN(year) || year < 0d || year >= 10000d)
        {
            return Error.Num;
        }

        if (year <= 1899d)
        {
            year += 1900d;
        }

        var months = Math.Truncate(monthArg);

        // AddMonths takes an int; bound the magnitude so the cast can't overflow (10000 years of months is
        // already well past DateTime's range, which the try/catch below would reject anyway).
        if (double.IsNaN(months) || Math.Abs(months) > 120000d)
        {
            return Error.Num;
        }

        var days = Math.Truncate(dayArg);

        if (double.IsNaN(days))
        {
            return Error.Num;
        }

        try
        {
            var date = new DateTime((int)year, 1, 1).AddMonths((int)months - 1).AddDays(days - 1d);
            var result = FromDateTime(date);

            if (result < 0d)
            {
                return Error.Num;
            }

            serial = result;
            return null;
        }
        catch (ArgumentOutOfRangeException)
        {
            return Error.Num;
        }
    }

    /// <summary>
    /// The time of day of a serial, rounded to the nearest whole second in <c>[0, 86400)</c>. Rounding (not
    /// truncation) matches Excel: <c>SECOND(TIME(10,30,45))</c> is 45 even though the reconstructed double
    /// can land a hair below 45 in IEEE-754. A serial that rounds up to a full day wraps back to 0.
    /// </summary>
    public static int TimeOfDaySeconds(double serial)
    {
        var fraction = serial - Math.Floor(serial);
        var seconds = (int)Math.Round(fraction * 86400d, MidpointRounding.AwayFromZero);
        return seconds >= 86400 ? 0 : seconds;
    }
}
