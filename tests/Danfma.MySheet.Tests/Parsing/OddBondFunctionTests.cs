using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Odd first/last coupon bonds (ODDFPRICE/ODDFYIELD/ODDLPRICE/ODDLYIELD), cross-checked against the
/// <see cref="Financial"/> oracle using the Microsoft documentation examples. The formulas were fuzzed
/// against the oracle across bases and frequencies.
/// </summary>
public class OddBondFunctionTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task OddFPrice_MatchesExcel()
    {
        var expected = Financial.OddFPrice(
            new DateTime(2008, 11, 11),
            new DateTime(2021, 3, 1),
            new DateTime(2008, 10, 15),
            new DateTime(2009, 3, 1),
            0.0785,
            0.0625,
            100,
            Frequency.SemiAnnual,
            DayCountBasis.UsPsa30_360
        );
        await Assert
            .That(
                Num(
                    Calc(
                        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)"
                    )
                )
            )
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task OddFYield_MatchesExcel()
    {
        var expected = Financial.OddFYield(
            new DateTime(2008, 11, 11),
            new DateTime(2021, 3, 1),
            new DateTime(2008, 10, 15),
            new DateTime(2009, 3, 1),
            0.0575,
            84.5,
            100,
            Frequency.SemiAnnual,
            DayCountBasis.UsPsa30_360
        );
        await Assert
            .That(
                Num(
                    Calc(
                        "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)"
                    )
                )
            )
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task OddFPrice_MisalignedFirstCoupon_ReturnsNum()
    {
        // firstCoupon 2009-04-01 does not line up with the schedule stepping back from maturity 2021-03-01.
        await Assert
            .That(
                Calc(
                    "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,4,1),0.0785,0.0625,100,2,0)"
                )
            )
            .IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task OddLPrice_MatchesExcel()
    {
        var expected = Financial.OddLPrice(
            new DateTime(2008, 2, 7),
            new DateTime(2008, 6, 15),
            new DateTime(2007, 10, 15),
            0.0375,
            0.0405,
            100,
            Frequency.SemiAnnual,
            DayCountBasis.UsPsa30_360
        );
        await Assert
            .That(
                Num(
                    Calc(
                        "=ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)"
                    )
                )
            )
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task OddLYield_MatchesExcel()
    {
        var expected = Financial.OddLYield(
            new DateTime(2008, 4, 20),
            new DateTime(2008, 6, 15),
            new DateTime(2007, 12, 24),
            0.0375,
            99.875,
            100,
            Frequency.SemiAnnual,
            DayCountBasis.UsPsa30_360
        );
        await Assert
            .That(
                Num(
                    Calc(
                        "=ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)"
                    )
                )
            )
            .IsEqualTo(expected)
            .Within(1e-6);
    }

    [Test]
    public async Task OddLPrice_SettlementNotAfterLastInterest_ReturnsNum()
    {
        await Assert
            .That(
                Calc(
                    "=ODDLPRICE(DATE(2007,9,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)"
                )
            )
            .IsEqualTo(ErrorValue.Number);
    }
}
