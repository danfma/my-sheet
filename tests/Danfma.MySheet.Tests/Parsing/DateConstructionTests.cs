using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Wave 5 — Date and time, construction/extraction family: DATE, TIME, DATEVALUE, TIMEVALUE, YEAR, MONTH,
// DAY, HOUR, MINUTE, SECOND. Golden values are quoted from the official support.microsoft.com pages
// (fetched 2026-07-02) and cited per test. Dates are serial doubles (OADate epoch 1899-12-30); inputs are
// built with DATE(...) so the assertions round-trip through the same DateSerial helper.
public class DateConstructionTests
{
    private const double Tolerance = 1e-7;

    private static object? Calc(string formula) =>
        ExpressionParser
            .Parse(formula, new Workbook().Sheets.Add("Sheet1"))
            .Evaluate(new Workbook())
            .AsObject();

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // --- DATE: year 0-1899 → +1900, and month/day overflow, from the official DATE page examples. ---

    [Test]
    public async Task Date_AddsYearBelow1900()
    {
        // support.microsoft.com DATE: DATE(108,1,2) → January 2, 2008 (1900+108).
        await Assert.That(Num(Calc("=YEAR(DATE(108,1,2))"))).IsEqualTo(2008d);
        await Assert.That(Num(Calc("=MONTH(DATE(108,1,2))"))).IsEqualTo(1d);
        await Assert.That(Num(Calc("=DAY(DATE(108,1,2))"))).IsEqualTo(2d);
    }

    [Test]
    public async Task Date_OverflowsMonthAndDayLikeExcel()
    {
        // DATE page: DATE(2008,14,2) → 2/2/2009; DATE(2008,-3,2) → 9/2/2007;
        // DATE(2008,1,35) → 2/4/2008; DATE(2008,1,-15) → 12/16/2007.
        await Assert.That(Calc("=DATE(2008,14,2)=DATE(2009,2,2)") as bool?).IsTrue();
        await Assert.That(Calc("=DATE(2008,-3,2)=DATE(2007,9,2)") as bool?).IsTrue();
        await Assert.That(Calc("=DATE(2008,1,35)=DATE(2008,2,4)") as bool?).IsTrue();
        await Assert.That(Calc("=DATE(2008,1,-15)=DATE(2007,12,16)") as bool?).IsTrue();
    }

    [Test]
    public async Task Date_OutOfRangeYearIsNum()
    {
        // year < 0 or ≥ 10000 → #NUM! (DATE page: valid year range 0..9999).
        await Assert.That(Calc("=DATE(-1,1,1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=DATE(10000,1,1)")).IsEqualTo(ErrorValue.Number);
    }

    // --- TIME: component roll-over from the official TIME page examples. ---

    [Test]
    public async Task Time_MatchesTheGoldenFractions()
    {
        // TIME page: TIME(12,0,0)=0.5; TIME(16,48,10)=0.7001157.
        await Assert.That(Num(Calc("=TIME(12,0,0)"))).IsEqualTo(0.5d).Within(Tolerance);
        await Assert.That(Num(Calc("=TIME(16,48,10)"))).IsEqualTo(0.7001157d).Within(Tolerance);
    }

    [Test]
    public async Task Time_RollsOverComponents()
    {
        // TIME page: TIME(27,0,0)=TIME(3,0,0)=.125; TIME(0,750,0)=TIME(12,30,0)=.520833 (rounded display);
        // TIME(0,0,2000)=TIME(0,33,20)=.023148 (rounded). Asserted against the exact seconds-of-day fraction
        // the documented time represents (the page rounds; the engine keeps full precision).
        await Assert.That(Num(Calc("=TIME(27,0,0)"))).IsEqualTo(0.125d).Within(Tolerance);
        await Assert.That(Num(Calc("=TIME(0,750,0)"))).IsEqualTo(45000d / 86400d).Within(Tolerance);
        await Assert.That(Num(Calc("=TIME(0,0,2000)"))).IsEqualTo(2000d / 86400d).Within(Tolerance);
    }

    [Test]
    public async Task Time_NegativeComponentIsNum()
    {
        // Components accept 0..32767; a negative one → #NUM! (derived from the documented range).
        await Assert.That(Calc("=TIME(-1,0,0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=TIME(0,-1,0)")).IsEqualTo(ErrorValue.Number);
    }

    // --- DATEVALUE / TIMEVALUE: text parsing, invariant formats. ---

    [Test]
    public async Task DateValue_ParsesTheDocumentedFormats()
    {
        // DATEVALUE page: "8/22/2011"→40777; "22-MAY-2011"→40685; "2011/02/23"→40597.
        // ISO "2011-02-23" is the same serial (documented ISO form).
        await Assert.That(Num(Calc("=DATEVALUE(\"8/22/2011\")"))).IsEqualTo(40777d);
        await Assert.That(Num(Calc("=DATEVALUE(\"22-MAY-2011\")"))).IsEqualTo(40685d);
        await Assert.That(Num(Calc("=DATEVALUE(\"2011/02/23\")"))).IsEqualTo(40597d);
        await Assert.That(Num(Calc("=DATEVALUE(\"2011-02-23\")"))).IsEqualTo(40597d);
    }

    [Test]
    public async Task DateValue_UnparseableIsValueError()
    {
        await Assert.That(Calc("=DATEVALUE(\"not a date\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task TimeValue_ParsesTimeFraction()
    {
        // TIMEVALUE page: "2:24 AM"→0.10; "22-Aug-2011 6:35 AM"→0.2743 (rounded display; date part
        // discarded). 6:35 AM is exactly 23700 seconds of the day.
        await Assert.That(Num(Calc("=TIMEVALUE(\"2:24 AM\")"))).IsEqualTo(0.1d).Within(Tolerance);
        await Assert
            .That(Num(Calc("=TIMEVALUE(\"22-Aug-2011 6:35 AM\")")))
            .IsEqualTo(23700d / 86400d)
            .Within(Tolerance);
    }

    // --- YEAR/MONTH/DAY and HOUR/MINUTE/SECOND. ---

    [Test]
    public async Task YearMonthDay_ExtractCalendarParts()
    {
        // MONTH page: 15-Apr-11 → 4; DAY page: 15-Apr-11 → 15; YEAR page returns the calendar year.
        await Assert.That(Num(Calc("=YEAR(DATE(2023,7,5))"))).IsEqualTo(2023d);
        await Assert.That(Num(Calc("=MONTH(DATE(2011,4,15))"))).IsEqualTo(4d);
        await Assert.That(Num(Calc("=DAY(DATE(2011,4,15))"))).IsEqualTo(15d);
    }

    [Test]
    public async Task HourMinuteSecond_ExtractTimeParts()
    {
        // HOUR page: HOUR(0.75)=18. SECOND/MINUTE/HOUR round-trip a TIME to the nearest second
        // (Excel rounds, so SECOND(TIME(10,30,45))=45 despite IEEE-754 noise).
        await Assert.That(Num(Calc("=HOUR(0.75)"))).IsEqualTo(18d);
        await Assert.That(Num(Calc("=HOUR(TIME(10,30,45))"))).IsEqualTo(10d);
        await Assert.That(Num(Calc("=MINUTE(TIME(10,30,45))"))).IsEqualTo(30d);
        await Assert.That(Num(Calc("=SECOND(TIME(10,30,45))"))).IsEqualTo(45d);
    }

    [Test]
    public async Task Components_RejectNegativeSerial()
    {
        // Negative serials are out of range for the date functions → #NUM!.
        await Assert.That(Calc("=YEAR(-1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=HOUR(-0.5)")).IsEqualTo(ErrorValue.Number);
    }

    // --- Round-trip through DateSerial (round-trip requirement of the wave). ---

    [Test]
    public async Task Date_RoundTripsThroughComponents()
    {
        // DATE(2026,7,2) → serial → YEAR/MONTH/DAY recovers the components.
        await Assert.That(Num(Calc("=YEAR(DATE(2026,7,2))"))).IsEqualTo(2026d);
        await Assert.That(Num(Calc("=MONTH(DATE(2026,7,2))"))).IsEqualTo(7d);
        await Assert.That(Num(Calc("=DAY(DATE(2026,7,2))"))).IsEqualTo(2d);
    }

    // --- Documented limitation (§A6): serials 1..59 render one day behind Excel and serial 60 (Excel's
    // phantom 1900-02-29) is not representable — it collides with 1900-02-28. This is REGISTERED, not
    // corrected; real dates (serial ≥ 61 / 1900-03-01) are exact. ---

    [Test]
    public async Task Serial60_Is1900LeapYearLimitation()
    {
        // Excel would show serial 60 as 29-Feb-1900; here it is 28-Feb-1900 (via OADate), same as serial 59
        // renders 27-Feb. From serial 61 (1-Mar-1900) on, the mapping matches Excel exactly.
        await Assert.That(Num(Calc("=MONTH(60)"))).IsEqualTo(2d);
        await Assert.That(Num(Calc("=DAY(60)"))).IsEqualTo(28d);
        await Assert.That(Num(Calc("=MONTH(61)"))).IsEqualTo(3d);
        await Assert.That(Num(Calc("=DAY(61)"))).IsEqualTo(1d);
    }
}
