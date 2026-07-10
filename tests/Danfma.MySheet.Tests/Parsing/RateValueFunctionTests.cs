using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Wave-6 rate conversions, cumulative loan amounts and value helpers, cross-checked against the
/// <see cref="Financial"/> oracle.
/// </summary>
public class RateValueFunctionTests
{
    private const double Tolerance = 1e-9;

    private static object? Calc(string formula, params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = new NumberValue(value);
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task Effect_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=EFFECT(0.0525,4)")))
            .IsEqualTo(Financial.Effect(0.0525, 4))
            .Within(Tolerance);
    }

    [Test]
    public async Task Effect_InvalidArguments_ReturnNum()
    {
        await Assert.That(Calc("=EFFECT(0,4)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=EFFECT(0.05,0)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Nominal_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=NOMINAL(0.053543,4)")))
            .IsEqualTo(Financial.Nominal(0.053543, 4))
            .Within(Tolerance);
    }

    [Test]
    public async Task EffectNominal_RoundTrip()
    {
        var effect = Num(Calc("=EFFECT(0.06,12)"));
        await Assert
            .That(
                Num(
                    Calc(
                        $"=NOMINAL({effect.ToString(System.Globalization.CultureInfo.InvariantCulture)},12)"
                    )
                )
            )
            .IsEqualTo(0.06)
            .Within(1e-9);
    }

    [Test]
    public async Task Mirr_MatchesExcel()
    {
        var cells = new[]
        {
            ("A1", -120000.0),
            ("A2", 39000.0),
            ("A3", 30000.0),
            ("A4", 21000.0),
            ("A5", 37000.0),
            ("A6", 46000.0),
        };
        var expected = Financial.Mirr(
            new[] { -120000.0, 39000, 30000, 21000, 37000, 46000 },
            0.1,
            0.12
        );
        await Assert
            .That(Num(Calc("=MIRR(A1:A6,0.1,0.12)", cells)))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task Mirr_NoSignChange_ReturnsDivZero()
    {
        await Assert
            .That(Calc("=MIRR(A1:A3,0.1,0.12)", ("A1", 100), ("A2", 200), ("A3", 300)))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Rri_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=RRI(96,10000,11000)")))
            .IsEqualTo(Financial.Rri(96, 10000, 11000))
            .Within(Tolerance);
    }

    [Test]
    public async Task PDuration_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=PDURATION(0.025,2000,2200)")))
            .IsEqualTo(Financial.Pduration(0.025, 2000, 2200))
            .Within(Tolerance);
    }

    [Test]
    public async Task PDuration_InvalidArguments_ReturnNum()
    {
        await Assert.That(Calc("=PDURATION(0,2000,2200)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task ISPmt_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=ISPMT(0.1/12,1,36,8000000)")))
            .IsEqualTo(Financial.ISPmt(0.1 / 12, 1, 36, 8000000))
            .Within(1e-6);
    }

    [Test]
    public async Task CumIPmt_MatchesExcel()
    {
        var expected = Financial.CumIPmt(0.09 / 12, 360, 125000, 13, 24, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=CUMIPMT(0.09/12,360,125000,13,24,0)")))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task CumIPmt_FirstPayment_MatchesExcel()
    {
        var expected = Financial.CumIPmt(0.09 / 12, 360, 125000, 1, 1, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=CUMIPMT(0.09/12,360,125000,1,1,0)")))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task CumPrinc_MatchesExcel()
    {
        var expected = Financial.CumPrinc(0.09 / 12, 360, 125000, 13, 24, PaymentDue.EndOfPeriod);
        await Assert
            .That(Num(Calc("=CUMPRINC(0.09/12,360,125000,13,24,0)")))
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task CumIPmt_InvalidRange_ReturnsNum()
    {
        await Assert.That(Calc("=CUMIPMT(0.09/12,360,125000,0,24,0)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task FvSchedule_MatchesExcel()
    {
        var expected = Financial.FvSchedule(1, new[] { 0.09, 0.11, 0.1 });
        await Assert
            .That(Num(Calc("=FVSCHEDULE(1,A1:A3)", ("A1", 0.09), ("A2", 0.11), ("A3", 0.1))))
            .IsEqualTo(expected)
            .Within(Tolerance);
    }

    [Test]
    public async Task DollarDe_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=DOLLARDE(1.02,16)")))
            .IsEqualTo(Financial.DollarDe(1.02, 16))
            .Within(Tolerance);
        await Assert
            .That(Num(Calc("=DOLLARDE(1.1,32)")))
            .IsEqualTo(Financial.DollarDe(1.1, 32))
            .Within(Tolerance);
    }

    [Test]
    public async Task DollarFr_MatchesExcel()
    {
        await Assert
            .That(Num(Calc("=DOLLARFR(1.125,16)")))
            .IsEqualTo(Financial.DollarFr(1.125, 16))
            .Within(Tolerance);
    }

    [Test]
    public async Task Dollar_RoundTrip()
    {
        var de = Num(Calc("=DOLLARDE(1.02,16)"));
        await Assert
            .That(
                Num(
                    Calc(
                        $"=DOLLARFR({de.ToString(System.Globalization.CultureInfo.InvariantCulture)},16)"
                    )
                )
            )
            .IsEqualTo(1.02)
            .Within(1e-9);
    }

    [Test]
    public async Task DollarDe_ZeroFraction_ReturnsDivZero()
    {
        await Assert.That(Calc("=DOLLARDE(1.02,0)")).IsEqualTo(ErrorValue.DivByZero);
    }
}
