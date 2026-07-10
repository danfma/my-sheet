using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue; // TUnit também define um StringValue

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// MOD/QUOTIENT/SIGN/PI/PRODUCT/SUMSQ/MULTINOMIAL/SERIESSUM. Golden values from the Microsoft
/// support pages (fetched 2026-07-01); PI uses <see cref="Math.PI"/> as the oracle.
/// </summary>
public class ArithmeticTests
{
    private const double Tolerance = 1e-9;

    private static object? Calc(string formula, params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task Mod_TakesTheSignOfTheDivisor()
    {
        // support.microsoft.com MOD: MOD(3,2)=1; MOD(-3,2)=1; MOD(3,-2)=-1; MOD(-3,-2)=-1
        // ("The result has the same sign as divisor"; MOD(n,d) = n - d*INT(n/d)).
        await Assert.That(Num(Calc("=MOD(3,2)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=MOD(-3,2)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=MOD(3,-2)"))).IsEqualTo(-1.0);
        await Assert.That(Num(Calc("=MOD(-3,-2)"))).IsEqualTo(-1.0);
        await Assert.That(Num(Calc("=MOD(3.5,2)"))).IsEqualTo(1.5).Within(Tolerance);
    }

    [Test]
    public async Task Mod_ZeroDivisorIsDivZero()
    {
        await Assert.That(Calc("=MOD(3,0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Quotient_MatchesExcelDocs()
    {
        // support.microsoft.com QUOTIENT: QUOTIENT(5,2)=2; QUOTIENT(4.5,3.1)=1; QUOTIENT(-10,3)=-3.
        await Assert.That(Num(Calc("=QUOTIENT(5,2)"))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=QUOTIENT(4.5,3.1)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=QUOTIENT(-10,3)"))).IsEqualTo(-3.0);
    }

    [Test]
    public async Task Quotient_ZeroDenominatorIsDivZero()
    {
        await Assert.That(Calc("=QUOTIENT(5,0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Sign_ThreeWay()
    {
        // support.microsoft.com SIGN: SIGN(10)=1; SIGN(4-4)=0; SIGN(-0.00001)=-1.
        await Assert.That(Num(Calc("=SIGN(10)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=SIGN(4-4)"))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=SIGN(-0.00001)"))).IsEqualTo(-1.0);
    }

    [Test]
    public async Task Pi_MatchesMath()
    {
        await Assert.That(Num(Calc("=PI()"))).IsEqualTo(Math.PI);
    }

    [Test]
    public async Task Product_MultipliesRangeAndDirectArguments()
    {
        // support.microsoft.com PRODUCT (data 5, 15, 30): PRODUCT(range)=2250; PRODUCT(range,2)=4500.
        var cells = new (string, Expression)[]
        {
            ("A1", new NumberValue(5)),
            ("A2", new NumberValue(15)),
            ("A3", new NumberValue(30)),
        };

        await Assert.That(Num(Calc("=PRODUCT(A1:A3)", cells))).IsEqualTo(2250.0);
        await Assert.That(Num(Calc("=PRODUCT(A1:A3,2)", cells))).IsEqualTo(4500.0);
        await Assert.That(Num(Calc("=PRODUCT(2,3,4)"))).IsEqualTo(24.0);
    }

    [Test]
    public async Task Product_IgnoresReferencedTextButCountsDirectNumericText()
    {
        // Excel rule shared with SUM: text in referenced cells is ignored; numeric text passed
        // directly is coerced.
        await Assert
            .That(
                Num(
                    Calc(
                        "=PRODUCT(A1:A2)",
                        ("A1", new NumberValue(5)),
                        ("A2", new StringValue("x"))
                    )
                )
            )
            .IsEqualTo(5.0);
        await Assert.That(Num(Calc("=PRODUCT(\"2\",3)"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task SumSq_SumsSquares()
    {
        // support.microsoft.com SUMSQ: SUMSQ(3,4)=25.
        await Assert.That(Num(Calc("=SUMSQ(3,4)"))).IsEqualTo(25.0);
        await Assert
            .That(
                Num(Calc("=SUMSQ(A1:A2)", ("A1", new NumberValue(3)), ("A2", new NumberValue(4))))
            )
            .IsEqualTo(25.0);
    }

    [Test]
    public async Task Multinomial_MatchesExcelDocs()
    {
        // support.microsoft.com MULTINOMIAL: MULTINOMIAL(2,3,4)=1260 (362880/288).
        await Assert.That(Num(Calc("=MULTINOMIAL(2,3,4)"))).IsEqualTo(1260.0);
        await Assert.That(Num(Calc("=MULTINOMIAL(3)"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task Multinomial_NegativeIsNum()
    {
        await Assert.That(Calc("=MULTINOMIAL(2,-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task SeriesSum_MatchesExcelDocsCosineExample()
    {
        // support.microsoft.com SERIESSUM: x=PI()/4, n=0, m=2, coefficients {1, -1/2!, 1/4!, -1/6!}
        // → 0.707103 (approximation of cos(45°)).
        var cells = new (string, Expression)[]
        {
            ("A1", new NumberValue(0.785398163)),
            ("B1", new NumberValue(1)),
            ("B2", new NumberValue(-0.5)),
            ("B3", new NumberValue(0.041666667)),
            ("B4", new NumberValue(-0.001388889)),
        };

        await Assert
            .That(Num(Calc("=SERIESSUM(A1,0,2,B1:B4)", cells)))
            .IsEqualTo(0.707103)
            .Within(1e-6);
    }

    [Test]
    public async Task SeriesSum_NonNumericCoefficientIsValueError()
    {
        // support.microsoft.com SERIESSUM: "If any argument is nonnumeric, SERIESSUM returns the
        // #VALUE! error value."
        await Assert
            .That(
                Calc(
                    "=SERIESSUM(1,0,1,B1:B2)",
                    ("B1", new NumberValue(1)),
                    ("B2", new StringValue("x"))
                )
            )
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task ArgumentErrors_Propagate()
    {
        await Assert.That(Calc("=MOD(1/0,2)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SIGN(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }
}
