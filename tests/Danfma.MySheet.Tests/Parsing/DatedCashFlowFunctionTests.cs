using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// XNPV and XIRR (dated cash flows), cross-checked against the <see cref="Financial"/> oracle. Includes a
/// stiff, long, irregular schedule to guard the bracketing solver (the RATE/IRR lesson).
/// </summary>
public class DatedCashFlowFunctionTests
{
    private static (Workbook Workbook, Sheet Sheet) Build(double[] flows, DateTime[] dates)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        for (var i = 0; i < flows.Length; i++)
        {
            sheet[$"A{i + 1}"] = new NumberValue(flows[i]);
            sheet[$"B{i + 1}"] = new NumberValue(dates[i].ToOADate());
        }

        return (workbook, sheet);
    }

    private static double Eval(string formula, double[] flows, DateTime[] dates)
    {
        var (workbook, sheet) = Build(flows, dates);
        var result = ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
        return result is double d ? d : double.NaN;
    }

    private static object? EvalObject(string formula, double[] flows, DateTime[] dates)
    {
        var (workbook, sheet) = Build(flows, dates);
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static readonly double[] Flows = { -10000.0, 2750, 4250, 3250, 2750 };
    private static readonly DateTime[] Dates =
    {
        new(2008, 1, 1),
        new(2008, 3, 1),
        new(2008, 10, 30),
        new(2009, 2, 15),
        new(2009, 4, 1),
    };

    [Test]
    public async Task XNpv_MatchesExcel()
    {
        var expected = Financial.XNpv(0.09, Flows, Dates);
        await Assert
            .That(Eval("=XNPV(0.09,A1:A5,B1:B5)", Flows, Dates))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task XIrr_MatchesExcel()
    {
        var expected = Financial.XIrr(Flows, Dates);
        await Assert
            .That(Eval("=XIRR(A1:A5,B1:B5)", Flows, Dates))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task XIrr_AtSolution_XnpvIsZero()
    {
        var irr = Eval("=XIRR(A1:A5,B1:B5)", Flows, Dates);
        var xnpv = Financial.XNpv(irr, Flows, Dates);
        // The rate is solved to ~1e-7; with dXNPV/drate on the order of 1e4 here, that leaves an XNPV
        // residual near 1e-3 against cash flows of ~1e4 — effectively zero.
        await Assert.That(xnpv).IsEqualTo(0).Within(1e-2);
    }

    [Test]
    public async Task XIrr_StiffLongIrregular_MatchesExcel()
    {
        // 35-year horizon, tiny interim flows, big terminal payoff — the kind of stiff problem naive
        // Newton fails on. The bracketing solver must still match the oracle.
        double[] flows = { -100000.0, 5, 7, 500000 };
        DateTime[] dates =
        {
            new(1990, 1, 1),
            new(2000, 6, 3),
            new(2010, 11, 20),
            new(2025, 3, 15),
        };
        var expected = Financial.XIrr(flows, dates);
        await Assert
            .That(Eval("=XIRR(A1:A4,B1:B4)", flows, dates))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task XIrr_OutOfOrderDates_MatchesExcel()
    {
        // Dates out of order are allowed; the first date is the anchor.
        double[] flows = { -10000.0, 4000, 3000, 5000 };
        DateTime[] dates = { new(2020, 1, 1), new(2021, 6, 1), new(2020, 9, 1), new(2022, 1, 1) };
        var expected = Financial.XIrr(flows, dates);
        await Assert
            .That(Eval("=XIRR(A1:A4,B1:B4)", flows, dates))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task XIrr_NoSignChange_ReturnsNum()
    {
        double[] flows = { 100.0, 200, 300 };
        DateTime[] dates = { new(2020, 1, 1), new(2020, 6, 1), new(2021, 1, 1) };
        await Assert
            .That(EvalObject("=XIRR(A1:A3,B1:B3)", flows, dates))
            .IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task XNpv_DateBeforeAnchor_ReturnsNum()
    {
        double[] flows = { -10000.0, 4000, 5000 };
        DateTime[] dates = { new(2020, 6, 1), new(2020, 1, 1), new(2021, 1, 1) };
        await Assert
            .That(EvalObject("=XNPV(0.09,A1:A3,B1:B3)", flows, dates))
            .IsEqualTo(ErrorValue.Number);
    }
}
