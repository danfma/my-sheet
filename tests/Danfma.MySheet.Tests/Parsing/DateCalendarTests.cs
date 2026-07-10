using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Wave 5 — Date and time, calendar-arithmetic family: DAYS, DAYS360, EDATE, EOMONTH, WEEKDAY, WEEKNUM,
// ISOWEEKNUM, DATEDIF, YEARFRAC. Golden values are from the official support.microsoft.com pages (and
// MS-OI29500 for the DAYS360/YEARFRAC algorithms), fetched 2026-07-02 and cited per test. Cases the pages
// do not tabulate (European DAYS360, EDATE clamp, DATEDIF "M"/"YM"/"MD", YEARFRAC bases 2/4) are DERIVED
// from the documented rules and marked as such.
public class DateCalendarTests
{
    private const double Tolerance = 1e-8;

    private static object? Calc(string formula) =>
        ExpressionParser
            .Parse(formula, new Workbook().Sheets.Add("Sheet1"))
            .Evaluate(new Workbook())
            .AsObject();

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // --- DAYS: whole days between two dates (may be negative). ---

    [Test]
    public async Task Days_CountsWholeDays()
    {
        // DAYS page: DAYS("15-MAR-2021","1-FEB-2021")=42; DAYS(31-Dec-2021, 1-Jan-2021)=364.
        await Assert.That(Num(Calc("=DAYS(DATE(2021,3,15),DATE(2021,2,1))"))).IsEqualTo(42d);
        await Assert.That(Num(Calc("=DAYS(DATE(2021,12,31),DATE(2021,1,1))"))).IsEqualTo(364d);
        // End before start → negative (derived from the definition).
        await Assert.That(Num(Calc("=DAYS(DATE(2021,2,1),DATE(2021,3,15))"))).IsEqualTo(-42d);
    }

    // --- DAYS360: 30/360 day count, US (NASD) default and European. ---

    [Test]
    public async Task Days360_UsMethodMatchesTheDocs()
    {
        // DAYS360 page (US/NASD, method omitted): 1/30/2011→2/1/2011 = 1; 1/1/2011→12/31/2011 = 360;
        // 1/1/2011→2/1/2011 = 30.
        await Assert.That(Num(Calc("=DAYS360(DATE(2011,1,30),DATE(2011,2,1))"))).IsEqualTo(1d);
        await Assert.That(Num(Calc("=DAYS360(DATE(2011,1,1),DATE(2011,12,31))"))).IsEqualTo(360d);
        await Assert.That(Num(Calc("=DAYS360(DATE(2011,1,1),DATE(2011,2,1))"))).IsEqualTo(30d);
    }

    [Test]
    public async Task Days360_EuropeanMethodDropsDay31()
    {
        // DERIVED from the European rule (MS-OI29500): both endpoints' day-31 → 30, no next-month roll.
        // 1/1/2011→12/31/2011: (12-1)*30 + (30-1) = 359.
        await Assert
            .That(Num(Calc("=DAYS360(DATE(2011,1,1),DATE(2011,12,31),TRUE)")))
            .IsEqualTo(359d);
    }

    // --- EDATE / EOMONTH. ---

    [Test]
    public async Task EDate_ShiftsMonthsAndClampsMonthEnd()
    {
        // EDATE page (start 15-Jan-11): +1 → 15-Feb-11; -1 → 15-Dec-10; +2 → 15-Mar-11.
        await Assert.That(Calc("=EDATE(DATE(2011,1,15),1)=DATE(2011,2,15)") as bool?).IsTrue();
        await Assert.That(Calc("=EDATE(DATE(2011,1,15),-1)=DATE(2010,12,15)") as bool?).IsTrue();
        await Assert.That(Calc("=EDATE(DATE(2011,1,15),2)=DATE(2011,3,15)") as bool?).IsTrue();
        // DERIVED clamp (the page shows no clamp example): 31-Jan + 1 month → 28-Feb (2011 non-leap).
        await Assert.That(Calc("=EDATE(DATE(2011,1,31),1)=DATE(2011,2,28)") as bool?).IsTrue();
    }

    [Test]
    public async Task EoMonth_ReturnsMonthEnd()
    {
        // EOMONTH page (start 1-Jan-11): +1 → 2/28/2011; -3 → 10/31/2010.
        await Assert.That(Calc("=EOMONTH(DATE(2011,1,1),1)=DATE(2011,2,28)") as bool?).IsTrue();
        await Assert.That(Calc("=EOMONTH(DATE(2011,1,1),-3)=DATE(2010,10,31)") as bool?).IsTrue();
    }

    // --- WEEKDAY: every return_type against Thursday 2/14/2008 (the WEEKDAY page's example date). ---

    [Test]
    public async Task Weekday_MatchesEveryReturnType()
    {
        // WEEKDAY page: A2=2/14/2008 (Thursday). WEEKDAY(A2)=5, (A2,2)=4, (A2,3)=3.
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14))"))).IsEqualTo(5d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),2)"))).IsEqualTo(4d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),3)"))).IsEqualTo(3d);
        // Types 11..17 derived from the page's own table for a Thursday: 4,3,2,1,7,6,5.
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),11)"))).IsEqualTo(4d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),12)"))).IsEqualTo(3d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),13)"))).IsEqualTo(2d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),14)"))).IsEqualTo(1d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),15)"))).IsEqualTo(7d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),16)"))).IsEqualTo(6d);
        await Assert.That(Num(Calc("=WEEKDAY(DATE(2008,2,14),17)"))).IsEqualTo(5d);
    }

    [Test]
    public async Task Weekday_UnknownReturnTypeIsNum()
    {
        await Assert.That(Calc("=WEEKDAY(DATE(2008,2,14),4)")).IsEqualTo(ErrorValue.Number);
    }

    // --- WEEKNUM / ISOWEEKNUM against 3/9/2012 (the pages' shared example date). ---

    [Test]
    public async Task WeekNum_MatchesSystem1AndIso()
    {
        // WEEKNUM page: A2=3/9/2012 → WEEKNUM(A2)=10 (Sunday-start), WEEKNUM(A2,2)=11 (Monday-start).
        await Assert.That(Num(Calc("=WEEKNUM(DATE(2012,3,9))"))).IsEqualTo(10d);
        await Assert.That(Num(Calc("=WEEKNUM(DATE(2012,3,9),1)"))).IsEqualTo(10d);
        await Assert.That(Num(Calc("=WEEKNUM(DATE(2012,3,9),2)"))).IsEqualTo(11d);
        await Assert.That(Num(Calc("=WEEKNUM(DATE(2012,3,9),11)"))).IsEqualTo(11d);
        await Assert.That(Num(Calc("=WEEKNUM(DATE(2012,3,9),17)"))).IsEqualTo(10d);
        // Type 21 = System 2 = ISO 8601 → 10 (matches ISOWEEKNUM).
        await Assert.That(Num(Calc("=WEEKNUM(DATE(2012,3,9),21)"))).IsEqualTo(10d);
    }

    [Test]
    public async Task IsoWeekNum_MatchesTheDocsAndYearBoundaries()
    {
        // ISOWEEKNUM page: 3/9/2012 → 10.
        await Assert.That(Num(Calc("=ISOWEEKNUM(DATE(2012,3,9))"))).IsEqualTo(10d);
        // DERIVED ISO boundaries: 1-Jan-2016 (Friday) belongs to week 53 of 2015; 4-Jan-2016 (Monday)
        // starts week 1 (the week containing the year's first Thursday).
        await Assert.That(Num(Calc("=ISOWEEKNUM(DATE(2016,1,1))"))).IsEqualTo(53d);
        await Assert.That(Num(Calc("=ISOWEEKNUM(DATE(2016,1,4))"))).IsEqualTo(1d);
    }

    // --- DATEDIF: the 6 units. ---

    [Test]
    public async Task DateDif_MatchesTheDocumentedUnits()
    {
        // DATEDIF page golden: ("1/1/2001","1/1/2003","Y")=2; ("6/1/2001","8/15/2002","D")=440;
        // (…,"YD")=75.
        await Assert.That(Num(Calc("=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"Y\")"))).IsEqualTo(2d);
        await Assert
            .That(Num(Calc("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"D\")")))
            .IsEqualTo(440d);
        await Assert
            .That(Num(Calc("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YD\")")))
            .IsEqualTo(75d);
        // DERIVED from the definitions: complete months "M"=14; months ignoring years "YM"=2.
        await Assert
            .That(Num(Calc("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"M\")")))
            .IsEqualTo(14d);
        await Assert
            .That(Num(Calc("=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YM\")")))
            .IsEqualTo(2d);
    }

    [Test]
    public async Task DateDif_MdUnitAndStartAfterEnd()
    {
        // "MD" (day difference, months/years ignored) — officially warned as unreliable; here a clean case:
        // 1-Jan → 15-Mar, days 15-1 = 14 (DERIVED).
        await Assert
            .That(Num(Calc("=DATEDIF(DATE(2001,1,1),DATE(2001,3,15),\"MD\")")))
            .IsEqualTo(14d);
        // DATEDIF page: start_date > end_date → #NUM!.
        await Assert
            .That(Calc("=DATEDIF(DATE(2003,1,1),DATE(2001,1,1),\"Y\")"))
            .IsEqualTo(ErrorValue.Number);
    }

    // --- YEARFRAC: the 5 bases (0,1,3 are page golden; 2,4 derived from the same day counts). ---

    [Test]
    public async Task YearFrac_MatchesTheFiveBases()
    {
        // YEARFRAC page (A2=1/1/2012, A3=7/30/2012): basis 0 (omitted)=0.58055556; basis 1=0.57650273
        // (366-day leap basis); basis 3=0.57808219 (365-day).
        await Assert
            .That(Num(Calc("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))")))
            .IsEqualTo(0.58055556d)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)")))
            .IsEqualTo(0.57650273d)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)")))
            .IsEqualTo(0.57808219d)
            .Within(Tolerance);
        // DERIVED from the same actual/30-360 day counts: basis 2 = 211/360; basis 4 = 209/360.
        await Assert
            .That(Num(Calc("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)")))
            .IsEqualTo(211d / 360d)
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),4)")))
            .IsEqualTo(209d / 360d)
            .Within(Tolerance);
    }

    [Test]
    public async Task YearFrac_InvalidBasisIsNum()
    {
        await Assert
            .That(Calc("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),5)"))
            .IsEqualTo(ErrorValue.Number);
    }
}
