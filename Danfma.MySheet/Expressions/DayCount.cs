namespace Danfma.MySheet.Expressions;

/// <summary>
/// Day-count conventions behind <c>YEARFRAC</c> — the five bases Excel documents (MS-OI29500 §18.17.7.352):
/// 0 = US (NASD) 30/360, 1 = actual/actual, 2 = actual/360, 3 = actual/365, 4 = European 30/360. Kept as a
/// standalone testable helper because the wave-6 bond functions (ACCRINT, PRICE, YIELD, COUP*, …) share the
/// exact same day counting. All methods assume <paramref name="start"/> ≤ <paramref name="end"/>; callers
/// order the pair first.
///
/// Known limitation, inherited from Excel and documented by Microsoft itself: basis 0 "may return an
/// incorrect result when the start_date is the last day in February". This matches the fictitious 30-Feb
/// adjustment and is reproduced, not corrected.
/// </summary>
internal static class DayCount
{
    /// <summary>YEARFRAC's day-count fraction for the given basis (0..4). Throws for an unknown basis; the
    /// function layer validates and maps that to <c>#NUM!</c> before calling.</summary>
    public static double YearFraction(DateTime start, DateTime end, int basis) => basis switch
    {
        0 => Nasd360Days(start, end) / 360d,
        1 => ActualDays(start, end) / AverageYearLength(start, end),
        2 => ActualDays(start, end) / 360d,
        3 => ActualDays(start, end) / 365d,
        4 => Euro360Days(start, end) / 360d,
        _ => throw new ArgumentOutOfRangeException(nameof(basis)),
    };

    /// <summary>Actual calendar days between the two dates.</summary>
    public static int ActualDays(DateTime start, DateTime end) => (end.Date - start.Date).Days;

    /// <summary>
    /// US (NASD) 30/360 day count used by YEARFRAC basis 0 (MS-OI29500): a February end-of-month is treated
    /// as day 30, and any day-31 is dropped to 30 (the end only when the start already sits on day 30).
    /// </summary>
    public static int Nasd360Days(DateTime start, DateTime end)
    {
        int d1 = start.Day, d2 = end.Day;

        // Last day of February on either endpoint is pulled to a nominal 30 (the "30-Feb" adjustment).
        var startIsFebEnd = start.Month == 2 && d1 == DateTime.DaysInMonth(start.Year, 2);

        if (startIsFebEnd)
        {
            if (end.Month == 2 && d2 == DateTime.DaysInMonth(end.Year, 2))
            {
                d2 = 30;
            }

            d1 = 30;
        }

        if (d1 == 31)
        {
            d1 = 30;
        }

        if (d2 == 31 && d1 == 30)
        {
            d2 = 30;
        }

        return (end.Year - start.Year) * 360 + (end.Month - start.Month) * 30 + (d2 - d1);
    }

    /// <summary>European 30/360 day count (basis 4): both endpoints' day-31 unconditionally becomes 30.</summary>
    public static int Euro360Days(DateTime start, DateTime end)
    {
        var d1 = start.Day == 31 ? 30 : start.Day;
        var d2 = end.Day == 31 ? 30 : end.Day;

        return (end.Year - start.Year) * 360 + (end.Month - start.Month) * 30 + (d2 - d1);
    }

    /// <summary>
    /// Denominator for basis 1 (actual/actual): the average length of every calendar year the range crosses
    /// (MS-OI29500 note d), so a range over 2011–2012 uses (365+366)/2. A single-year range just uses that
    /// year's length.
    /// </summary>
    private static double AverageYearLength(DateTime start, DateTime end)
    {
        if (start.Year == end.Year)
        {
            return DateTime.IsLeapYear(start.Year) ? 366d : 365d;
        }

        var total = 0;

        for (var year = start.Year; year <= end.Year; year++)
        {
            total += DateTime.IsLeapYear(year) ? 366 : 365;
        }

        return total / (double)(end.Year - start.Year + 1);
    }
}
