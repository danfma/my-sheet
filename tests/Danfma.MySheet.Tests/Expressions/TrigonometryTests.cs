using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The full trig family. Oracle for the pure math is .NET <see cref="Math"/>; Excel-specific
/// semantics (ATAN2 argument order, COT(0), the 2^27 bound of the 2013 additions) cite the
/// Microsoft support pages (fetched 2026-07-01).
/// </summary>
public class TrigonometryTests
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
    public async Task Sin_Cos_Tan_MatchMath()
    {
        await Assert.That(Num(Calc("=SIN(1)"))).IsEqualTo(Math.Sin(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=COS(1)"))).IsEqualTo(Math.Cos(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=TAN(1)"))).IsEqualTo(Math.Tan(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=SIN(PI()/6)"))).IsEqualTo(0.5).Within(Tolerance);
    }

    [Test]
    public async Task Cot_MatchesExcelDocs()
    {
        // support.microsoft.com COT: COT(30)=-0.156; COT(45)=0.617; COT(0)=#DIV/0!;
        // |number| >= 2^27 → #NUM!.
        await Assert.That(Num(Calc("=COT(30)"))).IsEqualTo(-0.156).Within(1e-3);
        await Assert.That(Num(Calc("=COT(45)"))).IsEqualTo(0.617).Within(1e-3);
        await Assert.That(Num(Calc("=COT(1)"))).IsEqualTo(Math.Cos(1) / Math.Sin(1)).Within(Tolerance);
        await Assert.That(Calc("=COT(0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=COT(134217728)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Sec_Csc_AreReciprocals()
    {
        await Assert.That(Num(Calc("=SEC(1)"))).IsEqualTo(1 / Math.Cos(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=CSC(1)"))).IsEqualTo(1 / Math.Sin(1)).Within(Tolerance);
        await Assert.That(Calc("=CSC(0)")).IsEqualTo(ErrorValue.DivByZero);
        // The 2013 trig additions share the documented 2^27 argument bound.
        await Assert.That(Calc("=SEC(134217728)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=CSC(-134217728)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Asin_Acos_Atan_MatchMath()
    {
        await Assert.That(Num(Calc("=ASIN(0.5)"))).IsEqualTo(Math.Asin(0.5)).Within(Tolerance);
        await Assert.That(Num(Calc("=ACOS(0.5)"))).IsEqualTo(Math.Acos(0.5)).Within(Tolerance);
        await Assert.That(Num(Calc("=ATAN(1)"))).IsEqualTo(Math.PI / 4).Within(Tolerance);
    }

    [Test]
    public async Task Asin_Acos_OutsideDomainIsNum()
    {
        await Assert.That(Calc("=ASIN(2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=ASIN(-1.1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=ACOS(2)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Atan2_UsesExcelArgumentOrder()
    {
        // support.microsoft.com ATAN2(x_num, y_num): ATAN2(1,1)=0.785398163;
        // ATAN2(-1,-1)=-2.35619449 — the INVERSE of .NET's Math.Atan2(y,x).
        await Assert.That(Num(Calc("=ATAN2(1,1)"))).IsEqualTo(0.785398163).Within(1e-9);
        await Assert.That(Num(Calc("=ATAN2(-1,-1)"))).IsEqualTo(-2.35619449).Within(1e-8);
        await Assert.That(Num(Calc("=ATAN2(1,2)"))).IsEqualTo(Math.Atan2(2, 1)).Within(Tolerance);
    }

    [Test]
    public async Task Atan2_OriginIsDivZero()
    {
        // support.microsoft.com ATAN2: "If both x_num and y_num are 0, ATAN2 returns #DIV/0!".
        await Assert.That(Calc("=ATAN2(0,0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Acot_MatchesExcelDocs()
    {
        // support.microsoft.com ACOT: ACOT(2)=0.4636; result lies in (0, pi).
        await Assert.That(Num(Calc("=ACOT(2)"))).IsEqualTo(0.4636).Within(1e-4);
        await Assert.That(Num(Calc("=ACOT(-2)"))).IsEqualTo(Math.PI / 2 - Math.Atan(-2)).Within(Tolerance);
        await Assert.That(Num(Calc("=ACOT(0)"))).IsEqualTo(Math.PI / 2).Within(Tolerance);
    }

    [Test]
    public async Task Hyperbolics_MatchMath()
    {
        await Assert.That(Num(Calc("=SINH(1)"))).IsEqualTo(Math.Sinh(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=COSH(1)"))).IsEqualTo(Math.Cosh(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=TANH(1)"))).IsEqualTo(Math.Tanh(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=COTH(2)"))).IsEqualTo(1 / Math.Tanh(2)).Within(Tolerance);
        await Assert.That(Num(Calc("=SECH(1)"))).IsEqualTo(1 / Math.Cosh(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=CSCH(1)"))).IsEqualTo(1 / Math.Sinh(1)).Within(Tolerance);
    }

    [Test]
    public async Task Hyperbolics_DomainErrors()
    {
        await Assert.That(Calc("=COTH(0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=CSCH(0)")).IsEqualTo(ErrorValue.DivByZero);
        // sinh/cosh overflow the double range around |x| > 710 → #NUM! like Excel.
        await Assert.That(Calc("=SINH(711)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=COSH(711)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task InverseHyperbolics_MatchMath()
    {
        await Assert.That(Num(Calc("=ASINH(1)"))).IsEqualTo(Math.Asinh(1)).Within(Tolerance);
        await Assert.That(Num(Calc("=ACOSH(2)"))).IsEqualTo(Math.Acosh(2)).Within(Tolerance);
        await Assert.That(Num(Calc("=ATANH(0.5)"))).IsEqualTo(Math.Atanh(0.5)).Within(Tolerance);
        // support.microsoft.com ACOTH: ACOTH(6)=0.168 (= atanh(1/6)).
        await Assert.That(Num(Calc("=ACOTH(6)"))).IsEqualTo(0.168).Within(1e-3);
        await Assert.That(Num(Calc("=ACOTH(6)"))).IsEqualTo(Math.Atanh(1.0 / 6)).Within(Tolerance);
    }

    [Test]
    public async Task InverseHyperbolics_DomainErrors()
    {
        await Assert.That(Calc("=ACOSH(0.5)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=ATANH(1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=ATANH(-1)")).IsEqualTo(ErrorValue.Number);
        // support.microsoft.com ACOTH: "The absolute value of Number must be greater than 1."
        await Assert.That(Calc("=ACOTH(0.5)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=ACOTH(1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Degrees_Radians_Convert()
    {
        // support.microsoft.com DEGREES: DEGREES(PI())=180. RADIANS(270)=4.712389 (3π/2).
        await Assert.That(Num(Calc("=DEGREES(PI())"))).IsEqualTo(180.0).Within(Tolerance);
        await Assert.That(Num(Calc("=RADIANS(270)"))).IsEqualTo(4.712389).Within(1e-6);
        await Assert.That(Num(Calc("=RADIANS(180)"))).IsEqualTo(Math.PI).Within(Tolerance);
    }

    [Test]
    public async Task ArgumentErrors_Propagate()
    {
        await Assert.That(Calc("=SIN(1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=ATAN2(\"abc\",1)")).IsEqualTo(ErrorValue.NotValue);
    }
}
