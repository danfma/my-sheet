namespace Danfma.MySheet.Expressions;

/// <summary>
/// Excel serial-date arithmetic for the date/time functions. A date IS a <c>double</c> serial (fiel ao
/// Excel; no dedicated value kind): the integer part counts days and the fraction is the time of day.
/// Conversion goes through .NET's OLE Automation date (<see cref="DateTime.FromOADate"/> /
/// <see cref="DateTime.ToOADate"/>), whose epoch is 1899-12-30 (serial 0) — this reproduces Excel's
/// fictitious 1900 leap year for serials on or after 61 (1900-03-01), which is the whole point of the
/// OA epoch.
///
/// Documented limitation (§A6 of the roadmap): serials 1..59 (Jan–Feb 1900) render one day AHEAD of Excel
/// (serial 1 is 1899-12-31 here, 1900-01-01 in Excel) and serial 60 (Excel's phantom 1900-02-29) is not
/// representable — <c>FromOADate(60)</c> yields 1900-02-28, colliding with serial 59. This is registered,
/// not corrected: real-world dates (≥ 1900-03-01) are exact and round-trip losslessly.
///
/// Functions treat a negative serial as out of range → <c>#NUM!</c>.
/// </summary>
internal static class DateSerial
{
    /// <summary>Largest representable serial: 9999-12-31 23:59:59 (the OADate upper bound).</summary>
    public static readonly double MaxSerial = new DateTime(9999, 12, 31, 23, 59, 59).ToOADate();

    /// <summary>
    /// serial → <see cref="DateTime"/> via OADate. Returns <see cref="Error.Num"/> when the serial is
    /// negative or beyond the representable range; <c>null</c> on success.
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
            dateTime = DateTime.FromOADate(serial);
            return null;
        }
        catch (ArgumentException)
        {
            dateTime = default;
            return Error.Num;
        }
    }

    /// <summary><see cref="DateTime"/> → serial (OADate). A date-only value yields an integer serial.</summary>
    public static double FromDateTime(DateTime dateTime) => dateTime.ToOADate();

    /// <summary>
    /// Builds a date serial from year/month/day with Excel's overflow rules (used by <c>DATE</c>, and the
    /// same normalization the other constructors rely on):
    /// <list type="bullet">
    /// <item>year 0..1899 → the value is added to 1900 (so <c>DATE(108,1,2)</c> is 2008-01-02);</item>
    /// <item>month outside 1..12 rolls the year (month 13 → January of the next year, month 0 → December of
    /// the previous year);</item>
    /// <item>day outside 1..(days in month) rolls the month (day 0 → last day of the previous month).</item>
    /// </list>
    /// year &lt; 0 or ≥ 10000, or any result outside 1900..9999, returns <see cref="Error.Num"/>.
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
            var result = date.ToOADate();

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
