using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Financial function tests. Expected values are cross-checked against
/// <see cref="Financial"/> (the ExcelFinancialFunctions package), a .NET port of Excel's
/// financial functions, so the asserts compare MySheet against an independent Excel oracle.
/// </summary>
public class FinancialFunctionTests
{
    private const double Tolerance = 1e-6;

    private static object? Calc(string formula, params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = new NumberValue(value);
        }

        return ExpressionParser.Parse(formula, sheet).Compute(workbook);
    }

    // Extracts the numeric result for tolerance comparison; a non-numeric result (e.g. an
    // ErrorValue) becomes NaN so the assert fails cleanly instead of throwing.
    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task Pmt_OrdinaryAnnuity_MatchesExcel()
    {
        var expected = Financial.Pmt(0.05, 10, -1000, 0, PaymentDue.EndOfPeriod);
        await Assert.That(Num(Calc("=PMT(0.05,10,-1000)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Pmt_PaymentAtBeginning_MatchesExcel()
    {
        var expected = Financial.Pmt(0.05, 10, -1000, 0, PaymentDue.BeginningOfPeriod);
        await Assert
            .That(Num(Calc("=PMT(0.05,10,-1000,0,1)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Pmt_ZeroRate_MatchesExcel()
    {
        var expected = Financial.Pmt(0, 10, -1000, 0, PaymentDue.EndOfPeriod);
        await Assert.That(Num(Calc("=PMT(0,10,-1000)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Pmt_PropagatesArgumentError()
    {
        await Assert.That(Calc("=PMT(1/0,10,-1000)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Pv_OrdinaryAnnuity_MatchesExcel()
    {
        var expected = Financial.Pv(0.05, 10, -100, -1000, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=PV(0.05,10,-100,-1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Pv_PaymentAtBeginning_MatchesExcel()
    {
        var expected = Financial.Pv(0.05, 10, -100, -1000, PaymentDue.BeginningOfPeriod);
        await Assert
            .That(Num(Calc("=PV(0.05,10,-100,-1000,1)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Pv_ZeroRate_MatchesExcel()
    {
        var expected = Financial.Pv(0, 10, -100, -1000, PaymentDue.EndOfPeriod);
        await Assert.That(Num(Calc("=PV(0,10,-100,-1000)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Pv_PropagatesArgumentError()
    {
        await Assert.That(Calc("=PV(1/0,10,-100)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Fv_OrdinaryAnnuity_MatchesExcel()
    {
        var expected = Financial.Fv(0.05, 10, -100, -1000, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=FV(0.05,10,-100,-1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Fv_PaymentAtBeginning_MatchesExcel()
    {
        var expected = Financial.Fv(0.05, 10, -100, -1000, PaymentDue.BeginningOfPeriod);
        await Assert
            .That(Num(Calc("=FV(0.05,10,-100,-1000,1)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Fv_ZeroRate_MatchesExcel()
    {
        var expected = Financial.Fv(0, 10, -100, -1000, PaymentDue.EndOfPeriod);
        await Assert.That(Num(Calc("=FV(0,10,-100,-1000)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Fv_PropagatesArgumentError()
    {
        await Assert.That(Calc("=FV(1/0,10,-100)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Nper_OrdinaryAnnuity_MatchesExcel()
    {
        var expected = Financial.NPer(0.05, -100, 1000, 0, PaymentDue.EndOfPeriod);
        await Assert.That(Num(Calc("=NPER(0.05,-100,1000)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Nper_ZeroRate_MatchesExcel()
    {
        var expected = Financial.NPer(0, -100, 1000, 0, PaymentDue.EndOfPeriod);
        await Assert.That(Num(Calc("=NPER(0,-100,1000)"))).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Nper_ImpossibleAmortization_ReturnsNum()
    {
        // k = pmt·(1+r·type)/r = 2000; (k − fv)/(k + pv) = 2000 / (2000 − 3000) = −2 → ln of a
        // non-positive number has no solution, so Excel returns #NUM!.
        await Assert.That(Calc("=NPER(0.05,100,-3000)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Nper_PropagatesArgumentError()
    {
        await Assert.That(Calc("=NPER(1/0,-100,1000)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Ipmt_FirstPeriod_MatchesExcel()
    {
        var expected = Financial.IPmt(0.05, 1, 10, -1000, 0, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=IPMT(0.05,1,10,-1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Ipmt_MiddlePeriod_MatchesExcel()
    {
        var expected = Financial.IPmt(0.05, 5, 10, -1000, 0, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=IPMT(0.05,5,10,-1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Ipmt_PaymentAtBeginning_MatchesExcel()
    {
        var expected = Financial.IPmt(0.05, 2, 10, -1000, 0, PaymentDue.BeginningOfPeriod);
        await Assert
            .That(Num(Calc("=IPMT(0.05,2,10,-1000,0,1)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Ipmt_PeriodOutOfRange_ReturnsNum()
    {
        await Assert.That(Calc("=IPMT(0.05,0,10,-1000)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=IPMT(0.05,11,10,-1000)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Ipmt_PropagatesArgumentError()
    {
        await Assert.That(Calc("=IPMT(1/0,1,10,-1000)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Ppmt_FirstPeriod_MatchesExcel()
    {
        var expected = Financial.PPmt(0.05, 1, 10, -1000, 0, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=PPMT(0.05,1,10,-1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Ppmt_MiddlePeriod_MatchesExcel()
    {
        var expected = Financial.PPmt(0.05, 5, 10, -1000, 0, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=PPMT(0.05,5,10,-1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Ppmt_PeriodOutOfRange_ReturnsNum()
    {
        await Assert.That(Calc("=PPMT(0.05,11,10,-1000)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task IpmtAndPpmt_SumToPayment()
    {
        // For any period, IPMT + PPMT = PMT (interest + principal = total payment).
        var pmt = Num(Calc("=PMT(0.05,10,-1000)"));
        var ipmt = Num(Calc("=IPMT(0.05,4,10,-1000)"));
        var ppmt = Num(Calc("=PPMT(0.05,4,10,-1000)"));
        await Assert.That(ipmt + ppmt).IsEqualTo(pmt).Within(Tolerance);
    }

    [Test]
    public async Task Npv_DirectArguments_MatchesExcel()
    {
        var expected = Financial.Npv(0.1, [-10000.0, 3000, 4200, 6800]);
        await Assert
            .That(Num(Calc("=NPV(0.1,-10000,3000,4200,6800)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Npv_RangeArgument_MatchesExcel()
    {
        var expected = Financial.Npv(0.1, [-10000.0, 3000, 4200, 6800]);
        var actual = Num(
            Calc("=NPV(0.1,A1:A4)", ("A1", -10000), ("A2", 3000), ("A3", 4200), ("A4", 6800))
        );
        await Assert.That(actual).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Npv_RateOfNegativeOne_ReturnsDivByZero()
    {
        await Assert.That(Calc("=NPV(-1,100,200)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Npv_PropagatesArgumentError()
    {
        await Assert.That(Calc("=NPV(0.1,1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task ClosedFormFunctions_SurviveSaveAndLoad()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=PMT(0.05,10,-1000)", sheet);
        sheet["A2"] = ExpressionParser.Parse("=PV(0.05,10,-100,-1000)", sheet);
        sheet["A3"] = ExpressionParser.Parse("=FV(0.05,10,-100,-1000)", sheet);
        sheet["A4"] = ExpressionParser.Parse("=NPER(0.05,-100,1000)", sheet);
        sheet["A5"] = ExpressionParser.Parse("=IPMT(0.05,1,10,-1000)", sheet);
        sheet["A6"] = ExpressionParser.Parse("=PPMT(0.05,1,10,-1000)", sheet);
        sheet["A7"] = ExpressionParser.Parse("=NPV(0.1,-10000,3000,4200,6800)", sheet);

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path);
            var loaded = Workbook.Load(path);
            var loadedSheet = loaded["Sheet1"];

            await Assert
                .That(Num(loadedSheet["A1"].Compute(loaded)))
                .IsEqualTo(Financial.Pmt(0.05, 10, -1000, 0, PaymentDue.EndOfPeriod))
                .Within(Tolerance);
            await Assert
                .That(Num(loadedSheet["A4"].Compute(loaded)))
                .IsEqualTo(Financial.NPer(0.05, -100, 1000, 0, PaymentDue.EndOfPeriod))
                .Within(Tolerance);
            await Assert
                .That(Num(loadedSheet["A7"].Compute(loaded)))
                .IsEqualTo(Financial.Npv(0.1, [-10000.0, 3000, 4200, 6800]))
                .Within(Tolerance);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Rate_InverseOfPmt_MatchesExcel()
    {
        var expected = Financial.Rate(10, -129.5046, 1000, 0, PaymentDue.EndOfPeriod, 0.1);
        await Assert
            .That(Num(Calc("=RATE(10,-129.5046,1000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Rate_LoanWithManyPeriods_MatchesExcel()
    {
        var expected = Financial.Rate(360, -600, 100000, 0, PaymentDue.EndOfPeriod, 0.1);
        await Assert
            .That(Num(Calc("=RATE(360,-600,100000)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Rate_WithExplicitGuess_MatchesExcel()
    {
        var expected = Financial.Rate(10, -129.5046, 1000, 0, PaymentDue.EndOfPeriod, 0.2);
        await Assert
            .That(Num(Calc("=RATE(10,-129.5046,1000,0,0,0.2)")))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Rate_NoRealSolution_ReturnsNum()
    {
        // 100·(1+r)² + 100 = 0 has no real root, so the solver fails to converge → #NUM!.
        await Assert.That(Calc("=RATE(2,0,100,100)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Rate_PropagatesArgumentError()
    {
        await Assert.That(Calc("=RATE(1/0,-100,1000)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Irr_TypicalCashFlows_MatchesExcel()
    {
        var expected = Financial.Irr([-10000.0, 3000, 4200, 6800], 0.1);
        var actual = Num(
            Calc("=IRR(A1:A4)", ("A1", -10000), ("A2", 3000), ("A3", 4200), ("A4", 6800))
        );
        await Assert.That(actual).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Irr_WithExplicitGuess_MatchesExcel()
    {
        var expected = Financial.Irr([-10000.0, 3000, 4200, 6800], 0.2);
        var actual = Num(
            Calc("=IRR(A1:A4,0.2)", ("A1", -10000), ("A2", 3000), ("A3", 4200), ("A4", 6800))
        );
        await Assert.That(actual).IsEqualTo(expected).Within(Tolerance);
    }

    [Test]
    public async Task Irr_AtSolution_NpvIsZero()
    {
        // Independent cross-check: NPV (period-0 convention) evaluated at the IRR must be ~0.
        var irr = Num(
            Calc("=IRR(A1:A4)", ("A1", -10000), ("A2", 3000), ("A3", 4200), ("A4", 6800))
        );
        var npvAtIrr =
            -10000 + 3000 / (1 + irr) + 4200 / Math.Pow(1 + irr, 2) + 6800 / Math.Pow(1 + irr, 3);
        // The IRR is solved to ~1e-7 on the rate; with dNPV/drate ≈ -15000 here, that leaves an NPV
        // residual on the order of 1e-3 — effectively zero against cash flows of ~1e4.
        await Assert.That(npvAtIrr).IsEqualTo(0).Within(1e-2);
    }

    [Test]
    public async Task Irr_NoSignChange_ReturnsNum()
    {
        await Assert
            .That(Calc("=IRR(A1:A3)", ("A1", 100), ("A2", 200), ("A3", 300)))
            .IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Irr_PropagatesArgumentError()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(-10000);
        sheet["A2"] = ExpressionParser.Parse("=1/0", sheet);
        sheet["A3"] = new NumberValue(4200);
        sheet["A4"] = new NumberValue(6800);

        var result = ExpressionParser.Parse("=IRR(A1:A4)", sheet).Compute(workbook);
        await Assert.That(result).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task IterativeFunctions_SurviveSaveAndLoad()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(-10000);
        sheet["A2"] = new NumberValue(3000);
        sheet["A3"] = new NumberValue(4200);
        sheet["A4"] = new NumberValue(6800);
        sheet["B1"] = ExpressionParser.Parse("=RATE(10,-129.5046,1000)", sheet);
        sheet["B2"] = ExpressionParser.Parse("=IRR(A1:A4)", sheet);

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path);
            var loaded = Workbook.Load(path);
            var loadedSheet = loaded["Sheet1"];

            await Assert
                .That(Num(loadedSheet["B1"].Compute(loaded)))
                .IsEqualTo(Financial.Rate(10, -129.5046, 1000, 0, PaymentDue.EndOfPeriod, 0.1))
                .Within(Tolerance);
            await Assert
                .That(Num(loadedSheet["B2"].Compute(loaded)))
                .IsEqualTo(Financial.Irr([-10000.0, 3000, 4200, 6800], 0.1))
                .Within(Tolerance);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
