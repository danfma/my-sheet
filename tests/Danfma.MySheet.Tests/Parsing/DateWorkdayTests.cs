using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

// Wave 5 — Date and time, working-day family: NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL. Golden
// values from the official support.microsoft.com pages (fetched 2026-07-02), cited per test. Holiday lists
// are supplied as cell ranges (resolved through ArgumentFlattening). Weekend numbers 1-7/11-17 and the
// 7-char "0000011" mask follow the shared NETWORKDAYS.INTL/WORKDAY.INTL weekend table.
public class DateWorkdayTests
{
    private static object? Calc(string formula, params (string Id, string Formula)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, cellFormula) in cells)
        {
            sheet[id] = ExpressionParser.Parse(cellFormula, sheet);
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // Holiday cells shared by the NETWORKDAYS page example: A2 start, A3 end, A4:A6 holidays.
    private static readonly (string, string)[] ProjectDates =
    [
        ("A2", "=DATE(2012,10,1)"),
        ("A3", "=DATE(2013,3,1)"),
        ("A4", "=DATE(2012,11,22)"),
        ("A5", "=DATE(2012,12,4)"),
        ("A6", "=DATE(2013,1,21)"),
    ];

    [Test]
    public async Task NetworkDays_MatchesTheDocExamples()
    {
        // NETWORKDAYS page: =NETWORKDAYS(A2,A3)=110; =NETWORKDAYS(A2,A3,A4)=109;
        // =NETWORKDAYS(A2,A3,A4:A6)=107.
        await Assert.That(Num(Calc("=NETWORKDAYS(A2,A3)", ProjectDates))).IsEqualTo(110d);
        await Assert.That(Num(Calc("=NETWORKDAYS(A2,A3,A4)", ProjectDates))).IsEqualTo(109d);
        await Assert.That(Num(Calc("=NETWORKDAYS(A2,A3,A4:A6)", ProjectDates))).IsEqualTo(107d);
    }

    [Test]
    public async Task NetworkDaysIntl_MatchesTheDocExamples()
    {
        // NETWORKDAYS.INTL page:
        //   (2006,1,1)→(2006,1,31) = 22; (2006,2,28)→(2006,1,31) = -21 (end before start).
        await Assert
            .That(Num(Calc("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))")))
            .IsEqualTo(22d);
        await Assert
            .That(Num(Calc("=NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))")))
            .IsEqualTo(-21d);
    }

    [Test]
    public async Task NetworkDaysIntl_WeekendNumberAndStringMask()
    {
        // NETWORKDAYS.INTL page, holidays H2=1/2/2006 and H3=1/16/2006:
        //   weekend 7 (Fri/Sat): (2006,1,1)→(2006,2,1) = 22;
        //   weekend "0010001" (Wed+Sun): (2006,1,1)→(2006,2,1) = 20.
        (string, string)[] holidays = [("H2", "=DATE(2006,1,2)"), ("H3", "=DATE(2006,1,16)")];
        await Assert
            .That(Num(Calc("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,H2:H3)", holidays)))
            .IsEqualTo(22d);
        await Assert
            .That(
                Num(
                    Calc(
                        "=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"0010001\",H2:H3)",
                        holidays
                    )
                )
            )
            .IsEqualTo(20d);
        // "1111111" (every day a weekend) always returns 0 (NETWORKDAYS.INTL page).
        await Assert
            .That(Num(Calc("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"1111111\")")))
            .IsEqualTo(0d);
    }

    [Test]
    public async Task NetworkDaysIntl_InvalidWeekendArgs()
    {
        // Invalid weekend number → #NUM!; malformed weekend string → #VALUE!.
        await Assert
            .That(Calc("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),8)"))
            .IsEqualTo(ErrorValue.Number);
        await Assert
            .That(Calc("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),\"00X0011\")"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Workday_MatchesTheDocExamples()
    {
        // WORKDAY page: start 10/1/2008, 151 days → 4/30/2009; with holidays 11/26/2008,12/4/2008,1/21/2009
        // → 5/5/2009. Compared as dates (the page's serials render as dates).
        (string, string)[] holidays =
        [
            ("W4", "=DATE(2008,11,26)"),
            ("W5", "=DATE(2008,12,4)"),
            ("W6", "=DATE(2009,1,21)"),
        ];
        await Assert.That(Calc("=WORKDAY(DATE(2008,10,1),151)=DATE(2009,4,30)") as bool?).IsTrue();
        await Assert
            .That(Calc("=WORKDAY(DATE(2008,10,1),151,W4:W6)=DATE(2009,5,5)", holidays) as bool?)
            .IsTrue();
    }

    [Test]
    public async Task WorkdayIntl_MatchesTheDocExamples()
    {
        // WORKDAY.INTL page: (2012,1,1),90,11 (Sunday-only weekend) → serial 41013 (= 4/14/2012);
        //   (2012,1,1),30,17 (Saturday-only) → 2/5/2012; (2012,1,1),30,0 → #NUM! (invalid weekend number).
        await Assert.That(Num(Calc("=WORKDAY.INTL(DATE(2012,1,1),90,11)"))).IsEqualTo(41013d);
        await Assert
            .That(Calc("=WORKDAY.INTL(DATE(2012,1,1),30,17)=DATE(2012,2,5)") as bool?)
            .IsTrue();
        await Assert.That(Calc("=WORKDAY.INTL(DATE(2012,1,1),30,0)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task WorkdayIntl_AllWeekendIsNum()
    {
        // WORKDAY.INTL with every day off has no day to land on → #NUM! (unlike NETWORKDAYS.INTL → 0).
        await Assert
            .That(Calc("=WORKDAY.INTL(DATE(2012,1,1),30,\"1111111\")"))
            .IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Workday_NegativeDaysWalkBackward()
    {
        // Negative days go to a past date (WORKDAY page: "a negative value yields a past date").
        // From Wednesday 1/11/2012 back 3 working days → Friday 1/6/2012 (DERIVED, default Sat/Sun weekend).
        await Assert.That(Calc("=WORKDAY(DATE(2012,1,11),-3)=DATE(2012,1,6)") as bool?).IsTrue();
    }
}
