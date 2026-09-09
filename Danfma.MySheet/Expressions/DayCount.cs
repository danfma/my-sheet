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
/// adjustment and is reproduced, not corrected. What is NOT reproduced is MS-OI29500's matching rule for the
/// END date — Excel as measured has no end-of-February adjustment at all; see
/// <see cref="Nasd360Days(int, int, int, int, int, int)"/>.
/// </summary>
internal static class DayCount
{
    /// <summary>
    /// YEARFRAC's day-count fraction for the given basis (0..4), over Excel date SERIALS. Throws for an
    /// unknown basis; the function layer validates and maps that to <c>#NUM!</c> before calling, and has
    /// already range-checked both serials and put them in order.
    /// </summary>
    /// <remarks>
    /// Bases 1-3 count days by subtracting the serials, never the mapped <see cref="DateTime"/>s, so the
    /// phantom 1900-02-29 is counted (measured <c>YEARFRAC(1,61,1)</c> = 60/365, <c>(59,61,1)</c> = 2/365,
    /// <c>(1,60,2)</c> = 59/360). Bases 0 and 4 do the opposite: they read the Lotus calendar, in which
    /// serial 60 IS February 29. Both measured on Aspose.Cells 26.6.0, PLAIN entry.
    /// </remarks>
    public static double YearFraction(double startSerial, double endSerial, int basis)
    {
        var start = Math.Floor(startSerial);
        var end = Math.Floor(endSerial);

        switch (basis)
        {
            case 1:
                return (end - start) / AverageYearLength(YearOf(start), YearOf(end));
            case 2:
                return (end - start) / 360d;
            case 3:
                return (end - start) / 365d;
            case < 0 or > 4:
                throw new ArgumentOutOfRangeException(nameof(basis));
        }

        LotusParts(start, out var y1, out var m1, out var d1);
        LotusParts(end, out var y2, out var m2, out var d2);

        // MEASURED and not derivable: the phantom day is day 29 when it OPENS the span (which the Lotus parts
        // already say) and day 30 when it CLOSES it — YEARFRAC(60,61,0) = 2/360 and (60,61,4) = 2/360 against
        // (1,60,0) = (1,60,4) = 59/360 and (31,60,0) = (31,60,4) = 30/360. That is the mirror image of
        // DAYS360, which opens on 30 and closes on 29.
        if (DateSerial.IsPhantomFeb29(end))
        {
            d2 = 30;
        }

        return basis == 0
            ? Nasd360Days(y1, m1, d1, y2, m2, d2) / 360d
            : Euro360Days(y1, m1, d1, y2, m2, d2) / 360d;
    }

    /// <summary>Actual calendar days between the two dates.</summary>
    public static int ActualDays(DateTime start, DateTime end) => (end.Date - start.Date).Days;

    /// <summary>
    /// US (NASD) 30/360 day count used by YEARFRAC basis 0, as MEASURED on Aspose.Cells 26.6.0 (PLAIN
    /// entry): a day-31 becomes 30 on the start, and on the end only when the start already sits on day 30;
    /// a start on the last day of February becomes 30 AFTER that test, so it does not drag a day-31 end down
    /// with it. There is no end-of-February rule at all.
    /// </summary>
    /// <remarks>
    /// The two departures from MS-OI29500 §18.17.7.352 are measured, not chosen: an end on the last day of
    /// February is NOT pulled to 30 — <c>YEARFRAC(DATE(2024,2,29),DATE(2025,2,28),0)</c> = 358/360,
    /// <c>(DATE(2011,2,28),DATE(2012,2,29),0)</c> = 359/360, <c>(DATE(2012,2,29),DATE(2013,2,28),0)</c> =
    /// 358/360, all of which the Microsoft rule answers 1 — and the ORDER of the February pull is what makes
    /// <c>(DATE(2023,2,28),DATE(2023,3,31),0)</c> = 31/360 and
    /// <c>(DATE(2024,2,29),DATE(2024,5,31),0)</c> = 91/360 rather than 30/360 and 90/360.
    /// DAYS360 orders the same two steps the other way round; see <c>Days360.UsDays360</c>.
    /// </remarks>
    public static int Nasd360Days(int y1, int m1, int d1, int y2, int m2, int d2)
    {
        // February's length here is the GREGORIAN one, so serial 59 (1900-02-28) is a February end for
        // basis 0 and is not one for DAYS360 (measured YEARFRAC(59,61,0) = 1/360, DAYS360(59,61) = 3).
        var startIsFebEnd = m1 == 2 && d1 == DateTime.DaysInMonth(y1, 2);

        if (d1 == 31)
        {
            d1 = 30;
        }

        if (d2 == 31 && d1 == 30)
        {
            d2 = 30;
        }

        if (startIsFebEnd)
        {
            d1 = 30;
        }

        return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
    }

    /// <summary>European 30/360 day count (basis 4): both endpoints' day-31 unconditionally becomes 30.</summary>
    public static int Euro360Days(DateTime start, DateTime end) =>
        Euro360Days(start.Year, start.Month, start.Day, end.Year, end.Month, end.Day);

    /// <inheritdoc cref="Euro360Days(DateTime, DateTime)"/>
    public static int Euro360Days(int y1, int m1, int d1, int y2, int m2, int d2)
    {
        if (d1 == 31)
        {
            d1 = 30;
        }

        if (d2 == 31)
        {
            d2 = 30;
        }

        return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
    }

    /// <summary>
    /// The calendar year a serial belongs to for the basis-1 denominator. Deliberately the MAPPED year, not
    /// the day-zero rule's 1900: measured <c>YEARFRAC(0,1500,1)</c> = 1500/365.1667 averages the years
    /// <b>1899</b>..1904, while <c>YEARFRAC(1,1500,1)</c> = 1499/365.2 averages 1900..1904.
    /// </summary>
    private static int YearOf(double serial) => DateSerial.ToDateTimeUnchecked(serial).Year;

    /// <summary>
    /// The Lotus year/month/day of a serial — serial 60 is 1900/2/29 — for the 30/360 bases. The range was
    /// checked by the caller, so the conversion cannot fail here.
    /// </summary>
    private static void LotusParts(double serial, out int year, out int month, out int day) =>
        _ = DateSerial.TryGetCalendar(serial, phantomFeb29: true, out year, out month, out day);

    /// <summary>
    /// Denominator for basis 1 (actual/actual): the average length of every calendar year the range crosses
    /// (MS-OI29500 note d), so a range over 2011–2012 uses (365+366)/2. A single-year range just uses that
    /// year's length.
    /// </summary>
    private static double AverageYearLength(int startYear, int endYear)
    {
        if (startYear == endYear)
        {
            return DateTime.IsLeapYear(startYear) ? 366d : 365d;
        }

        var total = 0;

        for (var year = startYear; year <= endYear; year++)
        {
            total += DateTime.IsLeapYear(year) ? 366 : 365;
        }

        return total / (double)(endYear - startYear + 1);
    }
}
