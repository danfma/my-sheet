using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Phase 9 — Excel's date serial epoch. Acceptance pins for the 1900 window: serial 1 is 1900-01-01, serial 0
// is the day-zero "1900-01-00", serial 60 is Excel's phantom 1900-02-29, and every function that COUNTS
// across that window spans one serial more than the Gregorian calendar does. MySheet converts through .NET's
// OLE Automation date (epoch 1899-12-30), so exactly the serials [0, 61) differ; from serial 61 (1900-03-01)
// on the two systems agree numerically.
//
// Provenance: every expected value in this file was MEASURED on Aspose.Cells 26.6.0 — PLAIN cell entry (a
// formula assigned to a cell), not CSE — and not quoted from a documentation page, because no page produces
// these numbers. Aspose is the oracle for "Excel" under the master plan's P0 rule.
//
// The Guard_* tests are GREEN today and must STAY green. They are the rows the central epoch map ALONE
// regresses: day counts that agree only because the old epoch's off-by-one cancels between the two ends, and
// the working-day rows the shifted calendar would move. A Guard_* failure is a REGRESSION, never progress.
// They are deliberately kept in their own tests so that a RED pin failing first can never hide one.
public class DateEpochTests
{
    // The 30/360 and Actual/365 pins are exact ratios of small integers; 1e-10 is far tighter than any
    // divergence this phase is about and far looser than IEEE-754 noise on a single division.
    private const double RatioTolerance = 1e-10;

    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // The dated cash-flow pins read their flows from K6:K7 and their dates from K8:K9 / K10:K11.
        sheet["K6"] = new NumberValue(-100d);
        sheet["K7"] = new NumberValue(110d);
        sheet["K8"] = new NumberValue(1d);
        sheet["K9"] = new NumberValue(366d);
        sheet["K10"] = new NumberValue(59d);
        sheet["K11"] = new NumberValue(61d);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(string formula) => Calc(formula) is double d ? d : double.NaN;

    // --- DATE / DATEVALUE: the constructor's own epoch. ---

    [Test]
    public async Task Date_PlacesSerial1On1900January1()
    {
        await Assert.That(Num("=DATE(1900,1,1)")).IsEqualTo(1d);
        await Assert.That(Num("=DATE(1900,1,0)")).IsEqualTo(0d);
        await Assert.That(Num("=DATE(1900,2,28)")).IsEqualTo(59d);
        await Assert.That(Num("=DATE(1900,3,0)")).IsEqualTo(59d);
        await Assert.That(Num("=DATE(0,1,1)")).IsEqualTo(1d);
    }

    [Test]
    public async Task Date_BelowSerialZeroIsNum()
    {
        // FromComponents keeps its "< 0 → #NUM!" guard, which the shifted epoch turns into a rejection of
        // DATE(1900,1,-1) (serial -1) where today it answers 0.
        await Assert.That(Calc("=DATE(1900,1,-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task DateValue_PlacesSerial1On1900January1()
    {
        await Assert.That(Num("=DATEVALUE(\"1900-01-01\")")).IsEqualTo(1d);
    }

    [Test]
    public async Task DateValue_ParsesThePhantomFebruary29AsSerial60()
    {
        // The phantom day is reachable by text: Aspose parses all three spellings to 60.
        await Assert.That(Num("=DATEVALUE(\"1900-02-29\")")).IsEqualTo(60d);
        await Assert.That(Num("=DATEVALUE(\"2/29/1900\")")).IsEqualTo(60d);
        await Assert.That(Num("=DATEVALUE(\"29-Feb-1900\")")).IsEqualTo(60d);
    }

    [Test]
    public async Task DateValue_RejectsDatesBefore1900()
    {
        // 1899-12-31 has no serial once serial 1 is 1900-01-01.
        await Assert.That(Calc("=DATEVALUE(\"1899-12-31\")")).IsEqualTo(ErrorValue.NotValue);
    }

    // --- YEAR / MONTH / DAY: the DateTime map plus the day-zero rule. ---

    [Test]
    public async Task YearMonthDay_ReadSerial1As1900January1()
    {
        await Assert.That(Num("=YEAR(1)")).IsEqualTo(1900d);
        await Assert.That(Num("=MONTH(1)")).IsEqualTo(1d);
        await Assert.That(Num("=DAY(1)")).IsEqualTo(1d);
    }

    [Test]
    public async Task YearMonthDay_ReadSerialZeroAsTheDayZero()
    {
        // Any serial in [0, 1) is "1900-01-00": year 1900, month 1, day 0.
        await Assert.That(Num("=YEAR(0)")).IsEqualTo(1900d);
        await Assert.That(Num("=MONTH(0)")).IsEqualTo(1d);
        await Assert.That(Num("=DAY(0)")).IsEqualTo(0d);
        await Assert.That(Num("=DAY(0.999)")).IsEqualTo(0d);
    }

    [Test]
    public async Task Day_ReadsSerial59AsFebruary28()
    {
        await Assert.That(Num("=DAY(59)")).IsEqualTo(28d);
    }

    // --- WEEKDAY / WEEKNUM / ISOWEEKNUM. ---

    [Test]
    public async Task Weekday_CollapsesSerial60OntoSerial59()
    {
        // The Lotus weekday is ((floor(s) - 1) mod 7) + 1 with 60 collapsed onto 59, so serial 60 is the
        // same Tuesday as serial 59 — one day behind what the OA epoch answers.
        await Assert.That(Num("=WEEKDAY(60)")).IsEqualTo(3d);
        await Assert.That(Num("=WEEKDAY(60,2)")).IsEqualTo(2d);
        await Assert.That(Num("=WEEKDAY(60,3)")).IsEqualTo(1d);
        await Assert.That(Num("=WEEKDAY(60.5)")).IsEqualTo(3d);
    }

    [Test]
    public async Task WeekNum_PlacesSerial1InTheFirstWeekOf1900()
    {
        await Assert.That(Num("=WEEKNUM(1)")).IsEqualTo(1d);
        await Assert.That(Num("=WEEKNUM(0)")).IsEqualTo(53d);
        await Assert.That(Num("=WEEKNUM(1,21)")).IsEqualTo(1d);
        await Assert.That(Num("=ISOWEEKNUM(1)")).IsEqualTo(1d);
    }

    // --- TEXT: the Lotus calendar (the only place a Feb 29 1900 is PRINTED). ---

    [Test]
    public async Task Text_PrintsTheLotusCalendar()
    {
        await Assert.That(Calc("=TEXT(1,\"yyyy-mm-dd\")") as string).IsEqualTo("1900-01-01");
        await Assert.That(Calc("=TEXT(60,\"yyyy-mm-dd\")") as string).IsEqualTo("1900-02-29");
        await Assert.That(Calc("=TEXT(0,\"yyyy-mm-dd\")") as string).IsEqualTo("1900-01-00");
        await Assert
            .That(Calc("=TEXT(60.5,\"yyyy-mm-dd hh:mm\")") as string)
            .IsEqualTo("1900-02-29 12:00");
    }

    [Test]
    public async Task Text_RejectsANegativeSerial()
    {
        await Assert.That(Calc("=TEXT(-1,\"yyyy-mm-dd\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Text_NamesSerial60TheSameWeekdayAsSerial59()
    {
        // The printed day NUMBER is the phantom 29th, but the printed day NAME comes from the Lotus weekday,
        // which collapses 60 onto 59 — so "1900-02-28 Tuesday" pairs a 28 with 59's weekday.
        await Assert.That(Calc("=TEXT(60,\"dddd\")") as string).IsEqualTo("Tuesday");
        await Assert
            .That(Calc("=TEXT(60,\"yyyy-mm-dd dddd\")") as string)
            .IsEqualTo("1900-02-28 Tuesday");
    }

    // --- EDATE / EOMONTH: month arithmetic on the DateTime map. ---

    [Test]
    public async Task EDate_ShiftsMonthsOnTheAsposeMap()
    {
        await Assert.That(Num("=EDATE(31,1)")).IsEqualTo(59d);
        await Assert.That(Num("=EDATE(60,0)")).IsEqualTo(59d);
        await Assert.That(Num("=EDATE(0,2)")).IsEqualTo(59d);
    }

    [Test]
    public async Task EDate_BelowSerialZeroIsNum()
    {
        await Assert.That(Calc("=EDATE(0,-1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=EDATE(1,-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task EoMonth_LandsOnJanuary31OfSerial31()
    {
        await Assert.That(Num("=EOMONTH(1,0)")).IsEqualTo(31d);
        await Assert.That(Num("=EOMONTH(0,0)")).IsEqualTo(31d);
        await Assert.That(Num("=EOMONTH(60,0)")).IsEqualTo(59d);
        await Assert.That(Num("=EOMONTH(1,-1)")).IsEqualTo(0d);
    }

    [Test]
    public async Task EoMonth_BelowSerialZeroIsNum()
    {
        await Assert.That(Calc("=EOMONTH(0,-1)")).IsEqualTo(ErrorValue.Number);
    }

    // --- DAYS360 / YEARFRAC: the 30/360 arithmetic reads the Lotus calendar (February 1900 has 29 days). ---

    [Test]
    public async Task Days360_CountsOnTheLotusCalendar()
    {
        await Assert.That(Num("=DAYS360(1,61)")).IsEqualTo(60d);
        await Assert.That(Num("=DAYS360(59,61)")).IsEqualTo(3d);
        await Assert.That(Num("=DAYS360(59,60)")).IsEqualTo(1d);
        await Assert.That(Num("=DAYS360(1,60)")).IsEqualTo(58d);
        await Assert.That(Num("=DAYS360(31,60)")).IsEqualTo(29d);
        await Assert.That(Num("=DAYS360(60,61,TRUE)")).IsEqualTo(1d);
    }

    [Test]
    public async Task Days360_ClampsAFebruaryEndOnModernDates()
    {
        // Generic (not epoch): a start on the last day of February makes the end day 30 too, and an end on
        // the last day of February is NOT promoted to 30 when the start was not. Prerequisite for the rows
        // above, so it is pinned on modern dates as well.
        await Assert.That(Num("=DAYS360(DATE(2024,1,31),DATE(2024,2,29))")).IsEqualTo(29d);
        await Assert.That(Num("=DAYS360(DATE(2024,2,28),DATE(2024,2,29))")).IsEqualTo(1d);
    }

    [Test]
    public async Task YearFrac_CountsOnTheLotusCalendar()
    {
        await Assert.That(Num("=YEARFRAC(1,366,4)")).IsEqualTo(359d / 360d).Within(RatioTolerance);
        await Assert.That(Num("=YEARFRAC(59,61)")).IsEqualTo(1d / 360d).Within(RatioTolerance);
        await Assert.That(Num("=YEARFRAC(60,61)")).IsEqualTo(2d / 360d).Within(RatioTolerance);
        await Assert.That(Num("=YEARFRAC(59,60)")).IsEqualTo(0d).Within(RatioTolerance);
        await Assert.That(Num("=YEARFRAC(1,60)")).IsEqualTo(59d / 360d).Within(RatioTolerance);
        await Assert.That(Num("=YEARFRAC(31,60,4)")).IsEqualTo(30d / 360d).Within(RatioTolerance);
    }

    // --- NETWORKDAYS / WORKDAY: the working-day family walks the shifted calendar. ---

    [Test]
    public async Task NetworkDays_CountsOnTheShiftedCalendar()
    {
        await Assert.That(Num("=NETWORKDAYS(1,10)")).IsEqualTo(8d);
        await Assert.That(Num("=NETWORKDAYS(1,1)")).IsEqualTo(1d);
        await Assert.That(Num("=NETWORKDAYS(6,6)")).IsEqualTo(0d);
        await Assert.That(Num("=NETWORKDAYS(1,61)")).IsEqualTo(45d);
        await Assert.That(Num("=NETWORKDAYS(0,1)")).IsEqualTo(1d);
        await Assert.That(Num("=NETWORKDAYS.INTL(1,8,\"0000011\")")).IsEqualTo(6d);
    }

    [Test]
    public async Task Workday_StepsOnTheShiftedCalendar()
    {
        await Assert.That(Num("=WORKDAY(1,5)")).IsEqualTo(8d);
        await Assert.That(Num("=WORKDAY(6,5)")).IsEqualTo(12d);
        await Assert.That(Num("=WORKDAY(8,5)")).IsEqualTo(15d);
        await Assert.That(Num("=WORKDAY(1,10)")).IsEqualTo(15d);
        await Assert.That(Num("=WORKDAY(2,-1)")).IsEqualTo(1d);
        await Assert.That(Num("=WORKDAY(8,-1)")).IsEqualTo(5d);
    }

    [Test]
    public async Task Workday_SteppingBelowSerialZeroIsNum()
    {
        // Serial 1 is a Monday on the shifted calendar, so one working day back is serial 0 — which is not a
        // date, and Aspose answers #NUM! rather than the -1 the OA epoch produces.
        await Assert.That(Calc("=WORKDAY(1,-1)")).IsEqualTo(ErrorValue.Number);
    }

    // --- DATEDIF: calendar units on the map, day units on serial subtraction. ---

    [Test]
    public async Task DateDif_MeasuresOnTheAsposeMap()
    {
        await Assert.That(Num("=DATEDIF(1,61,\"md\")")).IsEqualTo(0d);
        await Assert.That(Num("=DATEDIF(60,61,\"md\")")).IsEqualTo(2d);
        await Assert.That(Num("=DATEDIF(1,366,\"y\")")).IsEqualTo(0d);
        await Assert.That(Num("=DATEDIF(29,60,\"m\")")).IsEqualTo(0d);
    }

    // --- The bond / coupon family reads the DateTime map. ---

    [Test]
    public async Task CouponAndAccrual_ReadTheAsposeMap()
    {
        await Assert.That(Num("=COUPPCD(1,366,2)")).IsEqualTo(0d);
        await Assert.That(Num("=COUPDAYBS(59,366,2)")).IsEqualTo(58d);
        // Aspose: 16.666666666666664.
        await Assert
            .That(Num("=ACCRINT(1,182,61,0.1,1000,2)"))
            .IsEqualTo(16.666666666666664d)
            .Within(1e-9);
    }

    // ================================ GUARDS — MUST STAY GREEN ================================
    // Everything below already matches Aspose. The central epoch map ALONE regresses these rows, so they are
    // the proof that the per-function work landed with it. A failure here is a regression, not progress.

    [Test]
    public async Task Guard_DateAlreadyRollsTheNonExistentFebruary29()
    {
        // DATE builds a proleptic-Gregorian date, so Feb 29 1900 rolls to Mar 1 = serial 61 on BOTH engines.
        // Serial 60 is reachable only by arithmetic and by DATEVALUE.
        await Assert.That(Num("=DATE(1900,2,29)")).IsEqualTo(61d);
    }

    [Test]
    public async Task Guard_WeekdayOfSerial1AlreadyMatches()
    {
        // The Lotus weekday of serial 1 IS what FromOADate(1).DayOfWeek gives, so this row matches today —
        // and the shifted map must not move it.
        await Assert.That(Num("=WEEKDAY(1)")).IsEqualTo(1d);
        await Assert.That(Calc("=TEXT(1,\"dddd\")") as string).IsEqualTo("Sunday");
    }

    [Test]
    public async Task Guard_DayCountsWhoseEpochErrorCancelsBetweenTheEnds()
    {
        // Both ends shift by the same day, so the SPAN is unchanged — as long as the span does not straddle
        // serial 60. These are the rows a naive map change breaks by shifting only one end.
        await Assert.That(Num("=DAYS(60,59)")).IsEqualTo(1d);
        await Assert.That(Num("=DATEDIF(1,61,\"d\")")).IsEqualTo(60d);
        await Assert.That(Num("=YEARFRAC(1,61,1)")).IsEqualTo(60d / 365d).Within(RatioTolerance);
        await Assert.That(Num("=NETWORKDAYS(59,61)")).IsEqualTo(3d);
        await Assert.That(Num("=DAYS360(DATE(2023,2,28),DATE(2023,3,1))")).IsEqualTo(1d);
    }

    [Test]
    public async Task Guard_WorkdayRowsTheShiftedCalendarMustNotMove()
    {
        // The working-day walk lands on the same serial on both calendars for these shapes. WORKDAY(1,0) is
        // the cheapest proof that the days == 0 shortcut converts through the central map: it returned 2 on
        // the prototype that changed the map while that branch still called DateTime.ToOADate directly.
        await Assert.That(Num("=WORKDAY(1,0)")).IsEqualTo(1d);
        await Assert.That(Num("=WORKDAY(59,1)")).IsEqualTo(60d);
        await Assert.That(Num("=WORKDAY(60,-1)")).IsEqualTo(59d);
        await Assert.That(Num("=WORKDAY(5,1)")).IsEqualTo(6d);
        await Assert.That(Num("=WORKDAY(6,1)")).IsEqualTo(9d);
        await Assert.That(Num("=WORKDAY(6,4)")).IsEqualTo(12d);
        await Assert.That(Num("=WORKDAY(13,1)")).IsEqualTo(16d);
    }

    [Test]
    public async Task Guard_DatedCashFlowsWhoseEpochErrorCancels()
    {
        // XNPV/XIRR discount on serial DIFFERENCES, so a uniform shift cancels. The map alone regresses the
        // first to 0.026 and XIRR to 0.1003 by shifting the two dates unequally across serial 60.
        await Assert.That(Num("=XNPV(0.1,K6:K7,K8:K9)")).IsEqualTo(0d).Within(1e-9);
        await Assert
            .That(Num("=XNPV(0.1,K6:K7,K10:K11)"))
            .IsEqualTo(9.942567766564366d)
            .Within(1e-9);
        // A root-finder: pinned with a tolerance, not to four digits. Aspose 0.10000000000000009,
        // MySheet today 0.09999990463256836.
        await Assert.That(Num("=XIRR(K6:K7,K8:K9)")).IsEqualTo(0.1d).Within(1e-6);
    }
}
