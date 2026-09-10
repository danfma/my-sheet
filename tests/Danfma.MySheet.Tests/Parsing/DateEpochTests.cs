using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Phase 9 — Excel's date serial epoch. Acceptance pins for the 1900 window: serial 1 is 1900-01-01, serial 0
// is the day-zero "1900-01-00", serial 60 is Excel's phantom 1900-02-29, and every function that COUNTS
// across that window spans one serial more than the Gregorian calendar does. Before this phase MySheet
// converted through .NET's OLE Automation date (epoch 1899-12-30), so exactly the serials [0, 61) differed;
// from serial 61 (1900-03-01) on the two systems agree numerically, which is why the map change moves no
// modern date.
//
// Provenance — this file holds TWO kinds of expectation, and the difference matters:
//
// (1) ORACLE PINS, the default here. Every ORACLE value quoted in this file was MEASURED on Aspose.Cells
//     26.6.0 on 2026-09-09 — PLAIN cell entry (a formula assigned to a cell), not CSE — and never quoted from
//     a documentation page, because no page produces these numbers. Aspose is the oracle for "Excel" under
//     the master plan's P0 rule.
// (2) WALK PINS, the working-day rows below serial 61 in Guard_WorkdayRowsTheRealCalendarWalkMustNotMove and
//     Guard_WorkdayRowsWhereTheOracleContradictsItselfMustNotMove. Their EXPECTED value is what MySheet's
//     real-calendar serial walk produces, NOT what Aspose answers, by CONTROLLER RULING under the USER RULING
//     of 2026-09-09: Aspose has no derivable rule there and contradicts itself, so there is nothing
//     consistent to pin. Every one of those rows records Aspose's measured answer beside the assertion, so
//     the size of the accepted deviation stays visible; a row commented "agree" is both kinds at once. The
//     sub-61 rows of NetworkDays_CountsSerialsOnTheRealCalendar and Workday_StepsSerialsOnTheRealCalendar are
//     walk pins too — they simply coincide with the oracle on every shape they use, as each says.
//
// The Guard_* tests are GREEN today and must STAY green. Most of them are the rows the central epoch map
// ALONE regresses: day counts that agree only because the old epoch's off-by-one cancels between the two
// ends, and the working-day rows the shifted epoch map would move. The two working-day guards also carry the
// accepted deviations of kind (2), which is why they may not be deleted — a deleted row is an invisible
// deviation. A Guard_* failure is a REGRESSION, never progress. They are deliberately kept in their own tests
// so that a RED pin failing first can never hide one.
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

        // The modern working-day pins read their holidays from column M. March 2024: 45362 = Monday the 11th
        // … 45366 = Friday the 15th, 45367 Saturday, 45369 Monday, 45370 Tuesday, 45371 Wednesday.
        sheet["M1"] = new NumberValue(45367d); // Saturday — a holiday that falls ON a weekend
        sheet["M2"] = new NumberValue(45369d); // Monday — a holiday that falls on a working day
        sheet["M3"] = new NumberValue(45370d); // Tuesday
        sheet["M4"] = new NumberValue(45371d); // Wednesday
        sheet["M6"] = new NumberValue(45363d); // Tuesday, inside the span 45362..45366

        // A holiday inside the 1900 window, for the NETWORKDAYS row Aspose contradicts itself on. MySheet has
        // no array-constant syntax, so a holiday argument has to come from a cell.
        sheet["J1"] = new NumberValue(59d); // 1900-02-28 — inside the span 58..62

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
        // WEEKDAY(59) is the row the collapse is measured AGAINST, so it is asserted rather than implied: the
        // whole claim is that 60 answers what 59 answers.
        await Assert.That(Num("=WEEKDAY(59)")).IsEqualTo(3d);
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
        await Assert.That(Num("=DAYS360(DATE(2024,1,16),DATE(2024,2,29))")).IsEqualTo(43d);
    }

    [Test]
    public async Task Days360_HasNoEndOfMonthRollAtAll()
    {
        // Generic (not epoch), and the other half of the rule the 1900 rows need: support.microsoft.com and
        // MS-OI29500 §18.17.7.79 both roll a month-end END to the 1st of the next month when the adjusted
        // start is below day 30. Aspose has no such rule — a 30-day month end simply stays day 30.
        await Assert.That(Num("=DAYS360(DATE(2011,1,1),DATE(2011,4,30))")).IsEqualTo(119d);
        await Assert.That(Num("=DAYS360(DATE(2011,1,15),DATE(2011,9,30))")).IsEqualTo(255d);
        await Assert.That(Num("=DAYS360(DATE(2011,2,1),DATE(2011,4,30))")).IsEqualTo(89d);
        // A day-31 end still drops to 30 when the adjusted start reached 30, so these two do not move.
        await Assert.That(Num("=DAYS360(DATE(2011,1,1),DATE(2011,12,31))")).IsEqualTo(360d);
        await Assert.That(Num("=DAYS360(DATE(2024,1,30),DATE(2024,3,31))")).IsEqualTo(60d);
    }

    [Test]
    public async Task Days360_PullsAFebruaryEndBeforeTestingTheEnd()
    {
        // DAYS360 and YEARFRAC basis 0 order the same two 30/360 steps differently, so these rows and the
        // YEARFRAC ones below deliberately disagree: here the February-end start becomes day 30 FIRST and
        // therefore drags the day-31 end down to 30 as well.
        await Assert.That(Num("=DAYS360(DATE(2023,2,28),DATE(2023,3,31))")).IsEqualTo(30d);
        await Assert.That(Num("=DAYS360(DATE(2024,2,29),DATE(2024,5,31))")).IsEqualTo(90d);
        // 2024-02-28 is not February's last day, so nothing is pulled and the day-31 end survives.
        await Assert.That(Num("=DAYS360(DATE(2024,2,28),DATE(2024,5,31))")).IsEqualTo(93d);
    }

    [Test]
    public async Task Days360_OverThePhantomDayAloneIsZero()
    {
        // The phantom day counts as day 30 when it opens a span and day 29 when it closes one, which would
        // make a span from it to itself -1. Aspose answers 0.
        await Assert.That(Num("=DAYS360(60,60)")).IsEqualTo(0d);
        await Assert.That(Num("=DAYS360(60.5,60.9)")).IsEqualTo(0d);
        // The two roles themselves, and a reversed pair, which stays negative.
        await Assert.That(Num("=DAYS360(60,61)")).IsEqualTo(1d);
        await Assert.That(Num("=DAYS360(61,60)")).IsEqualTo(-2d);
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

    [Test]
    public async Task YearFrac_Basis0HasNoEndOfFebruaryRule()
    {
        // Generic (not epoch), and a prerequisite for the 1900 rows: MS-OI29500 pulls an END on the last day
        // of February to a nominal 30 when the start is one too, which makes all four of these exactly 1.
        // Aspose has no such rule.
        await Assert
            .That(Num("=YEARFRAC(DATE(2024,2,29),DATE(2025,2,28),0)"))
            .IsEqualTo(358d / 360d)
            .Within(RatioTolerance);
        await Assert
            .That(Num("=YEARFRAC(DATE(2011,2,28),DATE(2012,2,29),0)"))
            .IsEqualTo(359d / 360d)
            .Within(RatioTolerance);
        await Assert
            .That(Num("=YEARFRAC(DATE(2023,2,28),DATE(2024,2,29),0)"))
            .IsEqualTo(359d / 360d)
            .Within(RatioTolerance);
        await Assert
            .That(Num("=YEARFRAC(DATE(2012,2,29),DATE(2013,2,28),0)"))
            .IsEqualTo(358d / 360d)
            .Within(RatioTolerance);
        // The same rule inside the 1900 window: serial 59 to serial 425 is 1900-02-28 to 1901-02-28.
        await Assert.That(Num("=YEARFRAC(59,425,0)")).IsEqualTo(358d / 360d).Within(RatioTolerance);
    }

    [Test]
    public async Task YearFrac_Basis0PullsAFebruaryEndAfterTestingTheEnd()
    {
        // The start-of-February pull SURVIVES (these are 1/360 and 31/360, not 3/360 and 33/360) but it runs
        // AFTER the day-31 end test, so unlike DAYS360 it does not drag a day-31 end down to 30.
        await Assert
            .That(Num("=YEARFRAC(DATE(2023,2,28),DATE(2023,3,1),0)"))
            .IsEqualTo(1d / 360d)
            .Within(RatioTolerance);
        await Assert
            .That(Num("=YEARFRAC(DATE(2023,2,28),DATE(2023,3,31),0)"))
            .IsEqualTo(31d / 360d)
            .Within(RatioTolerance);
        await Assert
            .That(Num("=YEARFRAC(DATE(2024,2,29),DATE(2024,5,31),0)"))
            .IsEqualTo(91d / 360d)
            .Within(RatioTolerance);
        // Reached through the day-31 start rule instead, the pull-to-30 DOES take the end with it.
        await Assert
            .That(Num("=YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),0)"))
            .IsEqualTo(60d / 360d)
            .Within(RatioTolerance);
    }

    // --- NETWORKDAYS / WORKDAY: the working-day family walks the REAL calendar on serials. ---
    //
    // The walk runs on SERIALS and reads each weekday off the serial through the central map, so Excel's
    // phantom serial 60 is a day of the walk (a working Wednesday, the same day serial 59 names) and a holiday
    // set keyed by serial can never drift. Below serial 61 the walk is the REAL calendar by CONTROLLER RULING
    // under the USER RULING of 2026-09-09, not Aspose's composite; every row in the two tests immediately
    // below happens to agree with Aspose anyway, and the rows that do not are re-pinned in
    // Guard_WorkdayRowsTheRealCalendarWalkMustNotMove with both numbers recorded.

    [Test]
    public async Task NetworkDays_CountsSerialsOnTheRealCalendar()
    {
        // Aspose (26.6.0, 2026-09-09, PLAIN) and the real-calendar serial walk agree on every row here.
        await Assert.That(Num("=NETWORKDAYS(1,10)")).IsEqualTo(8d);
        await Assert.That(Num("=NETWORKDAYS(1,1)")).IsEqualTo(1d);
        await Assert.That(Num("=NETWORKDAYS(6,6)")).IsEqualTo(0d);
        // The span straddles the phantom day, so it counts one serial more than the Gregorian calendar has
        // days: 44 on a DateTime walk, 45 on the serial walk.
        await Assert.That(Num("=NETWORKDAYS(1,61)")).IsEqualTo(45d);
        await Assert.That(Num("=NETWORKDAYS(0,1)")).IsEqualTo(1d);
        await Assert.That(Num("=NETWORKDAYS.INTL(1,8,\"0000011\")")).IsEqualTo(6d);
    }

    [Test]
    public async Task Workday_StepsSerialsOnTheRealCalendar()
    {
        // Every argument is a serial <= 60, so by the CONTROLLER RULING these pin the real-calendar walk — and
        // on these six shapes the walk reproduces Aspose exactly (26.6.0, 2026-09-09, PLAIN: 8, 12, 15, 15, 1,
        // 5), so the expectation and the oracle coincide and nothing had to be re-pinned.
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
        // Serial 1 is a Monday on the real calendar and serial 0 is the Sunday 1899-12-31, so the backward walk
        // skips 0 as a weekend and the next candidate is the negative serial -1, which is not a date: #NUM!,
        // where the OA epoch answered -2. Aspose agrees (WORKDAY(1,-1) = #NUM!, measured 26.6.0, PLAIN).
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

    [Test]
    public async Task DateDif_AnchorsTheDayUnitsOnTheWholeMonthsAndYears()
    {
        // "MD" and "YD" count serials from the start pushed forward by every WHOLE month (year) the span
        // contains, with the shift CLAMPED to the target month's last day. Inside the 1900 window that is
        // what spans the phantom day; outside it, it is what keeps a day-31 start from overshooting the end
        // (the "borrow the previous month's length" formula answers -1 to the first two rows).
        await Assert.That(Num("=DATEDIF(DATE(2024,1,31),DATE(2024,3,1),\"md\")")).IsEqualTo(1d);
        await Assert.That(Num("=DATEDIF(DATE(2023,1,31),DATE(2023,3,1),\"md\")")).IsEqualTo(1d);
        await Assert.That(Num("=DATEDIF(DATE(2024,1,31),DATE(2024,4,1),\"md\")")).IsEqualTo(1d);
        await Assert.That(Num("=DATEDIF(DATE(2024,1,31),DATE(2024,2,29),\"md\")")).IsEqualTo(29d);
        await Assert.That(Num("=DATEDIF(DATE(2024,2,29),DATE(2025,3,1),\"yd\")")).IsEqualTo(1d);
        await Assert.That(Num("=DATEDIF(DATE(2024,2,29),DATE(2025,2,28),\"yd\")")).IsEqualTo(365d);
        await Assert.That(Num("=DATEDIF(0,60,\"md\")")).IsEqualTo(29d);
        await Assert.That(Num("=DATEDIF(60,60,\"md\")")).IsEqualTo(1d);
        await Assert.That(Num("=DATEDIF(1,366,\"yd\")")).IsEqualTo(365d);
        await Assert.That(Num("=DATEDIF(31,60,\"yd\")")).IsEqualTo(29d);
    }

    [Test]
    public async Task DateDif_OrdersThePairOnTheSerials()
    {
        // Serials 59 and 60 map to the same 1900-02-28, so a DateTime comparison lets a reversed pair through
        // and answers -1. Aspose rejects it.
        await Assert.That(Calc("=DATEDIF(60,59,\"d\")")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=DATEDIF(61,60,\"d\")")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Num("=DATEDIF(0,0,\"d\")")).IsEqualTo(0d);
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
        // Straddling the phantom day is the case that fails if a count goes through a DateTime: 61 and 59 map
        // two days apart on serials and only one day apart through the collapsed map.
        await Assert.That(Num("=DAYS(61,59)")).IsEqualTo(2d);
        await Assert.That(Num("=DATEDIF(1,61,\"d\")")).IsEqualTo(60d);
        await Assert.That(Num("=YEARFRAC(1,61,1)")).IsEqualTo(60d / 365d).Within(RatioTolerance);
        await Assert.That(Num("=NETWORKDAYS(59,61)")).IsEqualTo(3d);
        await Assert.That(Num("=DAYS360(DATE(2023,2,28),DATE(2023,3,1))")).IsEqualTo(1d);
    }

    [Test]
    public async Task Guard_WorkdayRowsTheRealCalendarWalkMustNotMove()
    {
        // RE-PINNED by CONTROLLER RULING at Task 2 close, under the USER RULING of 2026-09-09: every argument
        // here is a serial <= 60, so these rows are NOT Aspose pins any more. They pin the REAL-calendar serial
        // walk, and each row records BOTH numbers — the walk's value (the expectation) and Aspose's answer (the
        // accepted deviation, MEASURED on Aspose.Cells 26.6.0 on 2026-09-09, PLAIN cell entry). No row is
        // deleted: the size of the exception has to stay visible. WorkdayMath.IsWorkingSerial carries the
        // evidence that Aspose has no derivable rule below serial 61.
        //
        // WORKDAY(1,0) is the cheapest proof that the days == 0 shortcut converts through the central map: it
        // returned 2 on the prototype that changed the map while that branch still called DateTime.ToOADate.
        await Assert.That(Num("=WORKDAY(1,0)")).IsEqualTo(1d); // walk 1, Aspose 1 — agree
        await Assert.That(Num("=WORKDAY(59,1)")).IsEqualTo(60d); // walk 60, Aspose 60 — agree; the phantom
        // serial 60 is a working Wednesday, so the walk must not skip it (a DateTime walk answers 61).
        await Assert.That(Num("=WORKDAY(60,-1)")).IsEqualTo(59d); // walk 59, Aspose 59 — agree
        // The four rows below are the exception itself: 1900-01-05 is a Friday on the real calendar and a
        // Thursday on Aspose's Lotus weekday, and Aspose's own forward rule is not a function of `days`
        // (WORKDAY(6,4) = WORKDAY(6,5) = 12), so no walk can hold both.
        await Assert.That(Num("=WORKDAY(5,1)")).IsEqualTo(8d); // walk 8, Aspose 6 — accepted deviation
        await Assert.That(Num("=WORKDAY(6,1)")).IsEqualTo(8d); // walk 8, Aspose 9 — accepted deviation
        await Assert.That(Num("=WORKDAY(6,4)")).IsEqualTo(11d); // walk 11, Aspose 12 — accepted deviation
        await Assert.That(Num("=WORKDAY(13,1)")).IsEqualTo(15d); // walk 15, Aspose 16 — accepted deviation
    }

    [Test]
    public async Task Guard_WorkdayRowsWhereTheOracleContradictsItselfMustNotMove()
    {
        // The rows the function reference quotes as EXAMPLES of the sub-61 divergence: the four Aspose
        // self-contradictions listed there (a result that is not a function of `days`, a landing day Aspose
        // itself calls a weekend, and a non-additive NETWORKDAYS). Like the guard above, these are MySheet's
        // answers and NOT Aspose pins — they exist so the doc's numbers are pinned rather than unpinned prose
        // and cannot drift silently. Aspose's column is MEASURED on Aspose.Cells 26.6.0 on 2026-09-09, PLAIN
        // cell entry.
        //
        // Aspose answers 64 here, and 64 is a Sunday on BOTH calendars (WEEKDAY(64) = 1, NETWORKDAYS(64,64) =
        // 0, TEXT(64,"dddd") = Sunday), so no working-day walk can land on it.
        await Assert.That(Num("=WORKDAY(58,4)")).IsEqualTo(62d); // walk 62, Aspose 64
        // A span that CROSSES serial 61 — the structural half of the divergence, and the half no rule can
        // explain away. The walk counts the phantom serial 60 as the working Wednesday it repeats, so it is
        // ADDITIVE over a split; Aspose is not. Both engines answer 1, 3 and 1 to the three parts below, which
        // sum to the walk's 5 — Aspose's 4 for the whole span is therefore NOT producible by any per-day
        // working/non-working verdict, and that is precisely why additivity fails on its side.
        await Assert.That(Num("=NETWORKDAYS(58,62)")).IsEqualTo(5d); // walk 5, Aspose 4
        await Assert.That(Num("=NETWORKDAYS(58,58)")).IsEqualTo(1d); // walk 1, Aspose 1 — agree
        await Assert.That(Num("=NETWORKDAYS(59,61)")).IsEqualTo(3d); // walk 3, Aspose 3 — agree
        await Assert.That(Num("=NETWORKDAYS(62,62)")).IsEqualTo(1d); // walk 1, Aspose 1 — agree
        // Under a one-day weekend Aspose gives the SAME day for the 5th and the 6th working day, so the walk
        // can only reproduce one of the two rows; it reproduces the second.
        await Assert.That(Num("=WORKDAY.INTL(1,5,\"1000000\")")).IsEqualTo(6d); // walk 6, Aspose 7
        await Assert.That(Num("=WORKDAY.INTL(1,6,\"1000000\")")).IsEqualTo(7d); // walk 7, Aspose 7 — agree
        // The same crossing span with the J1 holiday (serial 59) taken out of it: 1 + 2 + 1 = 4 on both
        // engines part by part, and Aspose again answers one less than its own parts for the whole span.
        await Assert.That(Num("=NETWORKDAYS(58,62,J1)")).IsEqualTo(4d); // walk 4, Aspose 3
        await Assert.That(Num("=NETWORKDAYS(59,61,J1)")).IsEqualTo(2d); // walk 2, Aspose 2 — agree
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

    // =========================== MODERN-SERIAL PINS — MUST NEVER MOVE ===========================
    // Serials >= 61 must match Aspose byte-identical, and that is where every real workbook lives: the user's
    // exception covers ONLY the January–February 1900 window. Below serial 61 the working-day walk answers the
    // real calendar; from serial 61 on the epoch map is the identity, so the same walk is the oracle's own
    // answer and these rows are the regression guard for the whole family.
    //
    // Every value below MEASURED on Aspose.Cells 26.6.0 on 2026-09-09 — PLAIN cell entry (a formula assigned
    // to a cell), not CSE. March 2024: 45362 = Monday the 11th, 45366 = Friday the 15th.

    [Test]
    public async Task Modern_WorkdayStepsMatchTheOracle()
    {
        // Monday start: zero, one, a week, four weeks, and the same backward.
        await Assert.That(Num("=WORKDAY(45362,0)")).IsEqualTo(45362d);
        await Assert.That(Num("=WORKDAY(45362,1)")).IsEqualTo(45363d);
        await Assert.That(Num("=WORKDAY(45362,5)")).IsEqualTo(45369d);
        await Assert.That(Num("=WORKDAY(45362,20)")).IsEqualTo(45390d);
        await Assert.That(Num("=WORKDAY(45362,-1)")).IsEqualTo(45359d);
        await Assert.That(Num("=WORKDAY(45362,-5)")).IsEqualTo(45355d);
        await Assert.That(Num("=WORKDAY(45362,-20)")).IsEqualTo(45334d);
        // Friday start: the +1 step has to cross the weekend, the -1 step must not.
        await Assert.That(Num("=WORKDAY(45366,0)")).IsEqualTo(45366d);
        await Assert.That(Num("=WORKDAY(45366,1)")).IsEqualTo(45369d);
        await Assert.That(Num("=WORKDAY(45366,5)")).IsEqualTo(45373d);
        await Assert.That(Num("=WORKDAY(45366,20)")).IsEqualTo(45394d);
        await Assert.That(Num("=WORKDAY(45366,-1)")).IsEqualTo(45365d);
        await Assert.That(Num("=WORKDAY(45366,-5)")).IsEqualTo(45359d);
        await Assert.That(Num("=WORKDAY(45366,-20)")).IsEqualTo(45338d);
        // Holidays: one that falls on a Saturday costs nothing (M1); one on a Monday costs a day (M2).
        await Assert.That(Num("=WORKDAY(45366,5,M1)")).IsEqualTo(45373d);
        await Assert.That(Num("=WORKDAY(45366,5,M2)")).IsEqualTo(45376d);
        await Assert.That(Num("=WORKDAY(45366,5,M3:M4)")).IsEqualTo(45377d);
        await Assert.That(Num("=WORKDAY(45366,-5,M3:M4)")).IsEqualTo(45359d);
        await Assert.That(Num("=WORKDAY(45369,5,M3)")).IsEqualTo(45377d);
        await Assert.That(Num("=WORKDAY(45369,10,M3:M4)")).IsEqualTo(45385d);
        await Assert.That(Num("=WORKDAY(45372,-3,M3:M4)")).IsEqualTo(45365d);
        await Assert.That(Num("=WORKDAY(45366,0,M1)")).IsEqualTo(45366d);
        // `days` truncates toward zero, so a fraction never buys an extra step in either direction.
        await Assert.That(Num("=WORKDAY(45362,1.9)")).IsEqualTo(45363d);
        await Assert.That(Num("=WORKDAY(45362,-1.9)")).IsEqualTo(45359d);
    }

    [Test]
    public async Task Modern_WorkdayIntlWeekendFormsMatchTheOracle()
    {
        // Weekend numbers: 1 Sat+Sun (the default), 2 Sun+Mon, 3 Mon+Tue, 7 Fri+Sat; 11..17 are single days
        // (11 Sunday, 14 Wednesday, 17 Saturday).
        await Assert.That(Num("=WORKDAY.INTL(45362,5)")).IsEqualTo(45369d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,1)")).IsEqualTo(45369d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,2)")).IsEqualTo(45367d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,3)")).IsEqualTo(45368d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,7)")).IsEqualTo(45369d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,11)")).IsEqualTo(45367d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,14)")).IsEqualTo(45368d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,17)")).IsEqualTo(45368d);
        await Assert.That(Num("=WORKDAY.INTL(45366,5,11)")).IsEqualTo(45372d);
        await Assert.That(Num("=WORKDAY.INTL(45366,5,17)")).IsEqualTo(45372d);
        // String masks, Monday→Sunday: the two-day default, a one-day Sunday mask, a one-day Monday mask and a
        // mid-week pair.
        await Assert.That(Num("=WORKDAY.INTL(45362,5,\"0000011\")")).IsEqualTo(45369d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,\"0000001\")")).IsEqualTo(45367d);
        await Assert.That(Num("=WORKDAY.INTL(45362,5,\"1000000\")")).IsEqualTo(45367d);
        await Assert.That(Num("=WORKDAY.INTL(45362,20,\"0011000\")")).IsEqualTo(45390d);
        await Assert.That(Num("=WORKDAY.INTL(45362,-5,\"0000011\")")).IsEqualTo(45355d);
        await Assert.That(Num("=WORKDAY.INTL(45362,0,\"0000011\")")).IsEqualTo(45362d);
        await Assert.That(Num("=WORKDAY.INTL(45366,5,\"0000011\",M3:M4)")).IsEqualTo(45377d);
        await Assert.That(Num("=WORKDAY.INTL(45366,5,11,M3:M4)")).IsEqualTo(45374d);
        // Weekend 17 is Saturday-only, so the Saturday holiday in M1 is a real working-day loss.
        await Assert.That(Num("=WORKDAY.INTL(45366,20,17,M1)")).IsEqualTo(45390d);
        // An all-weekend mask has no day to land on: #VALUE!, not #NUM! (M4 of the phase's verifier).
        await Assert
            .That(Calc("=WORKDAY.INTL(45366,5,\"1111111\")"))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Num("=WORKDAY.INTL(45366,0,\"1111111\")")).IsEqualTo(45366d);
        // An out-of-table weekend number is #NUM! whichever side of the table it falls on.
        await Assert.That(Calc("=WORKDAY.INTL(45366,5,0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=WORKDAY.INTL(45366,5,8)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Modern_NetworkDaysCountsMatchTheOracle()
    {
        await Assert.That(Num("=NETWORKDAYS(45362,45366)")).IsEqualTo(5d);
        await Assert.That(Num("=NETWORKDAYS(45362,45362)")).IsEqualTo(1d);
        await Assert.That(Num("=NETWORKDAYS(45367,45367)")).IsEqualTo(0d);
        await Assert.That(Num("=NETWORKDAYS(45366,45373)")).IsEqualTo(6d);
        await Assert.That(Num("=NETWORKDAYS(45366,45362)")).IsEqualTo(-5d);
        await Assert.That(Num("=NETWORKDAYS(45362,45376)")).IsEqualTo(11d);
        await Assert.That(Num("=NETWORKDAYS(45362,45391)")).IsEqualTo(22d);
        await Assert.That(Num("=NETWORKDAYS(45366,45386)")).IsEqualTo(15d);
        // A holiday on a Saturday is already excluded (M1); one on a Tuesday inside the span costs a day (M6);
        // a range holding one of each costs exactly one (M1:M2).
        await Assert.That(Num("=NETWORKDAYS(45362,45366,M1)")).IsEqualTo(5d);
        await Assert.That(Num("=NETWORKDAYS(45362,45366,M6)")).IsEqualTo(4d);
        await Assert.That(Num("=NETWORKDAYS(45369,45373,M3:M4)")).IsEqualTo(3d);
        await Assert.That(Num("=NETWORKDAYS(45362,45376,M1:M2)")).IsEqualTo(10d);
        // Both ends floor to their whole day, so a time of day never adds or drops one.
        await Assert.That(Num("=NETWORKDAYS(45362.75,45366.25)")).IsEqualTo(5d);
    }

    [Test]
    public async Task Modern_NetworkDaysIntlWeekendFormsMatchTheOracle()
    {
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45366)")).IsEqualTo(5d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,1)")).IsEqualTo(11d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,2)")).IsEqualTo(10d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,3)")).IsEqualTo(10d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,7)")).IsEqualTo(11d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,11)")).IsEqualTo(13d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,14)")).IsEqualTo(13d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,17)")).IsEqualTo(13d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,\"0000011\")")).IsEqualTo(11d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,\"0000001\")")).IsEqualTo(13d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,\"1000000\")")).IsEqualTo(12d);
        // Unlike WORKDAY.INTL, an all-weekend mask is a legitimate zero here.
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,\"1111111\")")).IsEqualTo(0d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,\"0000011\",M3:M4)")).IsEqualTo(9d);
        // Weekend 11 is Sunday-only, so the Saturday holiday in M1 does cost a day.
        await Assert.That(Num("=NETWORKDAYS.INTL(45362,45376,11,M1)")).IsEqualTo(12d);
        await Assert.That(Num("=NETWORKDAYS.INTL(45376,45362,\"0000011\")")).IsEqualTo(-11d);
        await Assert.That(Calc("=NETWORKDAYS.INTL(45362,45376,0)")).IsEqualTo(ErrorValue.Number);
    }
}
