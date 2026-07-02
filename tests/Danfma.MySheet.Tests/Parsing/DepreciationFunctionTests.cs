using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Wave-6 depreciation functions. Expected values come from the <see cref="Financial"/> oracle
/// (ExcelFinancialFunctions), an independent .NET port of Excel's financial functions.
/// </summary>
public class DepreciationFunctionTests
{
    private const double Tolerance = 1e-9;

    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task Sln_MatchesExcel()
    {
        await Assert.That(Num(Calc("=SLN(30000,7500,10)"))).IsEqualTo(Financial.Sln(30000, 7500, 10)).Within(Tolerance);
    }

    [Test]
    public async Task Sln_ZeroLife_ReturnsDivZero()
    {
        await Assert.That(Calc("=SLN(30000,7500,0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Syd_MatchesExcel()
    {
        await Assert.That(Num(Calc("=SYD(30000,7500,10,1)"))).IsEqualTo(Financial.Syd(30000, 7500, 10, 1)).Within(Tolerance);
        await Assert.That(Num(Calc("=SYD(30000,7500,10,10)"))).IsEqualTo(Financial.Syd(30000, 7500, 10, 10)).Within(Tolerance);
    }

    [Test]
    public async Task Syd_PeriodOutOfRange_ReturnsNum()
    {
        await Assert.That(Calc("=SYD(30000,7500,10,11)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Db_MatchesExcel()
    {
        // MS example: DB(1000000,100000,6,1,7) = first partial year.
        await Assert.That(Num(Calc("=DB(1000000,100000,6,1,7)"))).IsEqualTo(Financial.Db(1000000, 100000, 6, 1, 7)).Within(Tolerance);
        await Assert.That(Num(Calc("=DB(1000000,100000,6,2,7)"))).IsEqualTo(Financial.Db(1000000, 100000, 6, 2, 7)).Within(Tolerance);
        // month defaults to 12 when omitted.
        await Assert.That(Num(Calc("=DB(1000000,100000,6,2)"))).IsEqualTo(Financial.Db(1000000, 100000, 6, 2, 12)).Within(Tolerance);
    }

    [Test]
    public async Task Ddb_MatchesExcel()
    {
        await Assert.That(Num(Calc("=DDB(2400,300,10,1)"))).IsEqualTo(Financial.Ddb(2400, 300, 10, 1, 2)).Within(Tolerance);
        await Assert.That(Num(Calc("=DDB(2400,300,10,5,1.5)"))).IsEqualTo(Financial.Ddb(2400, 300, 10, 5, 1.5)).Within(Tolerance);
    }

    [Test]
    public async Task Vdb_MatchesExcel()
    {
        await Assert.That(Num(Calc("=VDB(2400,300,3650,0,1)"))).IsEqualTo(Financial.Vdb(2400, 300, 3650, 0, 1)).Within(Tolerance);
        await Assert.That(Num(Calc("=VDB(2400,300,10,0,0.875,1.5)"))).IsEqualTo(Financial.Vdb(2400, 300, 10, 0, 0.875, 1.5)).Within(Tolerance);
        // no_switch = TRUE keeps declining balance.
        await Assert.That(Num(Calc("=VDB(2400,300,10,6,10,2,TRUE)")))
            .IsEqualTo(Financial.Vdb(2400, 300, 10, 6, 10, 2, VdbSwitch.DontSwitchToStraightLine))
            .Within(Tolerance);
    }

    [Test]
    public async Task AmorLinc_MatchesExcel()
    {
        // basis 1 (actual/actual): AMORLINC(2400, 2008-08-19, 2008-12-31, 300, 1, 0.15, 1).
        var expected = Financial.AmorLinc(2400, new DateTime(2008, 8, 19), new DateTime(2008, 12, 31), 300, 1, 0.15, DayCountBasis.ActualActual);
        await Assert.That(Num(Calc("=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task AmorLinc_FirstPeriod_MatchesExcel()
    {
        var expected = Financial.AmorLinc(2400, new DateTime(2008, 8, 19), new DateTime(2008, 12, 31), 300, 0, 0.15, DayCountBasis.ActualActual);
        await Assert.That(Num(Calc("=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,0,0.15,1)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task AmorLinc_Actual360Basis_ReturnsNum()
    {
        await Assert.That(Calc("=AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,2)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task AmorDegrc_MatchesExcel()
    {
        var expected = Financial.AmorDegrc(2400, new DateTime(2008, 8, 19), new DateTime(2008, 12, 31), 300, 1, 0.15, DayCountBasis.ActualActual, true);
        await Assert.That(Num(Calc("=AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task AmorDegrc_FirstAndLaterPeriods_MatchExcel()
    {
        for (var period = 0; period <= 5; period++)
        {
            var expected = Financial.AmorDegrc(2400, new DateTime(2008, 8, 19), new DateTime(2008, 12, 31), 300, period, 0.15, DayCountBasis.ActualActual, true);
            await Assert.That(Num(Calc($"=AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,{period},0.15,1)"))).IsEqualTo(expected).Within(Tolerance);
        }
    }
}
