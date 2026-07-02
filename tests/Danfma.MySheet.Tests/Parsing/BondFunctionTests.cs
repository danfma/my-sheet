using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Wave-6 bond math: coupon schedule, PRICE/YIELD/DURATION/MDURATION, accrued interest, single-period
/// discount securities and T-bills. Expected values come from the <see cref="Financial"/> oracle. Bond day
/// counts and coupon dates were fuzzed against the oracle across the five bases and three frequencies.
/// </summary>
public class BondFunctionTests
{
    private const double Tolerance = 1e-9;

    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    private static readonly DateTime Settlement = new(2011, 1, 25);
    private static readonly DateTime Maturity = new(2015, 11, 15);

    // ---- coupon schedule (semi-annual, US 30/360) ----

    [Test]
    public async Task CoupPcd_MatchesExcel()
    {
        var expected = Financial.CoupPCD(Settlement, Maturity, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360).ToOADate();
        await Assert.That(Num(Calc("=COUPPCD(DATE(2011,1,25),DATE(2015,11,15),2)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task CoupNcd_MatchesExcel()
    {
        var expected = Financial.CoupNCD(Settlement, Maturity, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360).ToOADate();
        await Assert.That(Num(Calc("=COUPNCD(DATE(2011,1,25),DATE(2015,11,15),2)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task CoupNum_MatchesExcel()
    {
        var expected = Financial.CoupNum(Settlement, Maturity, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);
        await Assert.That(Num(Calc("=COUPNUM(DATE(2011,1,25),DATE(2015,11,15),2)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task CoupDays_AcrossBases_MatchExcel()
    {
        foreach (var basis in new[] { 0, 1, 2, 3, 4 })
        {
            var expected = Financial.CoupDays(Settlement, Maturity, Frequency.SemiAnnual, (DayCountBasis)basis);
            await Assert.That(Num(Calc($"=COUPDAYS(DATE(2011,1,25),DATE(2015,11,15),2,{basis})"))).IsEqualTo(expected).Within(Tolerance);
        }
    }

    [Test]
    public async Task CoupDayBs_AcrossBases_MatchExcel()
    {
        foreach (var basis in new[] { 0, 1, 2, 3, 4 })
        {
            var expected = Financial.CoupDaysBS(Settlement, Maturity, Frequency.SemiAnnual, (DayCountBasis)basis);
            await Assert.That(Num(Calc($"=COUPDAYBS(DATE(2011,1,25),DATE(2015,11,15),2,{basis})"))).IsEqualTo(expected).Within(Tolerance);
        }
    }

    [Test]
    public async Task CoupDaysNc_AcrossBases_MatchExcel()
    {
        foreach (var basis in new[] { 0, 1, 2, 3, 4 })
        {
            var expected = Financial.CoupDaysNC(Settlement, Maturity, Frequency.SemiAnnual, (DayCountBasis)basis);
            await Assert.That(Num(Calc($"=COUPDAYSNC(DATE(2011,1,25),DATE(2015,11,15),2,{basis})"))).IsEqualTo(expected).Within(Tolerance);
        }
    }

    [Test]
    public async Task CoupFunctions_SettlementNotBeforeMaturity_ReturnNum()
    {
        await Assert.That(Calc("=COUPNUM(DATE(2015,11,15),DATE(2011,1,25),2)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Coup_InvalidFrequency_ReturnsNum()
    {
        await Assert.That(Calc("=COUPDAYS(DATE(2011,1,25),DATE(2015,11,15),3)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Coup_InvalidBasis_ReturnsNum()
    {
        await Assert.That(Calc("=COUPDAYS(DATE(2011,1,25),DATE(2015,11,15),2,5)")).IsEqualTo(ErrorValue.Number);
    }

    // ---- PRICE / YIELD / DURATION / MDURATION (MS doc examples) ----

    [Test]
    public async Task Price_MatchesExcel()
    {
        // MS example: settlement 2008-02-15, maturity 2017-11-15, coupon 5.75%, yield 6.5%, freq 2, basis 0.
        var expected = Financial.Price(new DateTime(2008, 2, 15), new DateTime(2017, 11, 15), 0.0575, 0.065, 100, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);
        await Assert.That(Num(Calc("=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task Price_AllBases_MatchExcel()
    {
        foreach (var basis in new[] { 0, 1, 2, 3, 4 })
        {
            var expected = Financial.Price(Settlement, Maturity, 0.05, 0.06, 100, Frequency.SemiAnnual, (DayCountBasis)basis);
            await Assert.That(Num(Calc($"=PRICE(DATE(2011,1,25),DATE(2015,11,15),0.05,0.06,100,2,{basis})"))).IsEqualTo(expected).Within(1e-6);
        }
    }

    [Test]
    public async Task Yield_MatchesExcel()
    {
        // MS example: settlement 2008-02-15, maturity 2016-11-15, coupon 5.75%, price 95.04287, freq 2, basis 0.
        var expected = Financial.Yield(new DateTime(2008, 2, 15), new DateTime(2016, 11, 15), 0.0575, 95.04287, 100, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);
        await Assert.That(Num(Calc("=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task PriceYield_InverseAcrossBases()
    {
        // YIELD must invert PRICE for every basis.
        foreach (var basis in new[] { 0, 1, 2, 3, 4 })
        {
            var price = Num(Calc($"=PRICE(DATE(2011,1,25),DATE(2015,11,15),0.05,0.06,100,2,{basis})"));
            var yield = Num(Calc($"=YIELD(DATE(2011,1,25),DATE(2015,11,15),0.05,{price.ToString(System.Globalization.CultureInfo.InvariantCulture)},100,2,{basis})"));
            await Assert.That(yield).IsEqualTo(0.06).Within(1e-6);
        }
    }

    [Test]
    public async Task Price_LastPeriod_MatchesExcel()
    {
        // Settlement inside the final coupon period (n == 1) uses the simple-interest branch.
        var expected = Financial.Price(new DateTime(2017, 7, 1), new DateTime(2017, 11, 15), 0.0575, 0.065, 100, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);
        await Assert.That(Num(Calc("=PRICE(DATE(2017,7,1),DATE(2017,11,15),0.0575,0.065,100,2,0)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task Duration_MatchesExcel()
    {
        var expected = Financial.Duration(new DateTime(2008, 1, 1), new DateTime(2016, 1, 1), 0.08, 0.09, Frequency.SemiAnnual, DayCountBasis.ActualActual);
        await Assert.That(Num(Calc("=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task MDuration_MatchesExcel()
    {
        var expected = Financial.MDuration(new DateTime(2008, 1, 1), new DateTime(2016, 1, 1), 0.08, 0.09, Frequency.SemiAnnual, DayCountBasis.ActualActual);
        await Assert.That(Num(Calc("=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)"))).IsEqualTo(expected).Within(1e-6);
    }

    // ---- accrued interest ----

    [Test]
    public async Task AccrInt_MatchesExcel()
    {
        var expected = Financial.AccrInt(new DateTime(2008, 3, 1), new DateTime(2008, 8, 31), new DateTime(2008, 5, 1), 0.1, 1000, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);
        await Assert.That(Num(Calc("=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task AccrInt_MultiPeriodAllBases_MatchExcel()
    {
        // Issue well before settlement (accrual spans several coupon periods); first_interest >= settlement.
        foreach (var basis in new[] { 0, 1, 2, 3, 4 })
        {
            var expected = Financial.AccrInt(new DateTime(2008, 1, 1), new DateTime(2011, 7, 1), new DateTime(2010, 4, 15), 0.08, 1000, Frequency.SemiAnnual, (DayCountBasis)basis);
            await Assert.That(Num(Calc($"=ACCRINT(DATE(2008,1,1),DATE(2011,7,1),DATE(2010,4,15),0.08,1000,2,{basis})"))).IsEqualTo(expected).Within(1e-6);
        }
    }

    [Test]
    public async Task AccrIntM_MatchesExcel()
    {
        var expected = Financial.AccrIntM(new DateTime(2008, 4, 1), new DateTime(2008, 6, 15), 0.1, 1000, DayCountBasis.Actual365);
        await Assert.That(Num(Calc("=ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)"))).IsEqualTo(expected).Within(1e-6);
    }

    // ---- single-period discount / interest securities ----

    [Test]
    public async Task Disc_MatchesExcel()
    {
        var expected = Financial.Disc(new DateTime(2018, 1, 25), new DateTime(2018, 6, 15), 97.975, 100, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=DISC(DATE(2018,1,25),DATE(2018,6,15),97.975,100,2)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task IntRate_MatchesExcel()
    {
        var expected = Financial.IntRate(new DateTime(2008, 2, 15), new DateTime(2008, 5, 15), 1000000, 1014420, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task Received_MatchesExcel()
    {
        var expected = Financial.Received(new DateTime(2008, 2, 15), new DateTime(2008, 5, 15), 1000000, 0.0575, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task PriceDisc_MatchesExcel()
    {
        var expected = Financial.PriceDisc(new DateTime(2008, 2, 16), new DateTime(2008, 3, 1), 0.0525, 100, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task YieldDisc_MatchesExcel()
    {
        var expected = Financial.YieldDisc(new DateTime(2008, 2, 16), new DateTime(2008, 3, 1), 99.795, 100, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task PriceMat_MatchesExcel()
    {
        var expected = Financial.PriceMat(new DateTime(2008, 2, 15), new DateTime(2008, 4, 13), new DateTime(2007, 11, 11), 0.061, 0.061, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,2)"))).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    public async Task YieldMat_MatchesExcel()
    {
        var expected = Financial.YieldMat(new DateTime(2008, 3, 15), new DateTime(2008, 11, 3), new DateTime(2007, 11, 8), 0.0625, 100.0123, DayCountBasis.Actual360);
        await Assert.That(Num(Calc("=YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,2)"))).IsEqualTo(expected).Within(1e-6);
    }

    // ---- T-bills ----

    [Test]
    public async Task TBillEq_MatchesExcel()
    {
        var expected = Financial.TBillEq(new DateTime(2008, 3, 31), new DateTime(2008, 6, 1), 0.0914);
        await Assert.That(Num(Calc("=TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task TBillPrice_MatchesExcel()
    {
        var expected = Financial.TBillPrice(new DateTime(2008, 3, 31), new DateTime(2008, 6, 1), 0.09);
        await Assert.That(Num(Calc("=TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task TBillYield_MatchesExcel()
    {
        var expected = Financial.TBillYield(new DateTime(2008, 3, 31), new DateTime(2008, 6, 1), 98.45);
        await Assert.That(Num(Calc("=TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)"))).IsEqualTo(expected).Within(1e-9);
    }

    [Test]
    public async Task TBill_MaturityOverOneYear_ReturnsNum()
    {
        await Assert.That(Calc("=TBILLPRICE(DATE(2008,3,31),DATE(2010,6,1),0.09)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Price_SurvivesSaveAndLoad()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)", sheet);
        sheet["A2"] = ExpressionParser.Parse("=COUPNUM(DATE(2011,1,25),DATE(2015,11,15),2)", sheet);

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path);
            var loaded = Workbook.Load(path);
            var loadedSheet = loaded["Sheet1"];

            var expectedPrice = Financial.Price(new DateTime(2008, 2, 15), new DateTime(2017, 11, 15), 0.0575, 0.065, 100, Frequency.SemiAnnual, DayCountBasis.UsPsa30_360);
            await Assert.That(Num(loadedSheet["A1"].Evaluate(loaded).AsObject())).IsEqualTo(expectedPrice).Within(1e-6);
            await Assert.That(Num(loadedSheet["A2"].Evaluate(loaded).AsObject())).IsEqualTo(10).Within(Tolerance);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
