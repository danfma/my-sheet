using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// SQRT/POWER/EXP/LN/LOG/LOG10/SQRTPI. Oracles: .NET <see cref="Math"/> for the pure math; the
/// Excel-specific cases cite the Microsoft support pages (POWER, LOG, SQRTPI — fetched 2026-07-01).
/// </summary>
public class PowerAndLogTests
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
    public async Task Sqrt_MatchesMath()
    {
        await Assert.That(Num(Calc("=SQRT(16)"))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=SQRT(2)"))).IsEqualTo(Math.Sqrt(2)).Within(Tolerance);
        await Assert.That(Num(Calc("=SQRT(0)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Sqrt_NegativeIsNum()
    {
        await Assert.That(Calc("=SQRT(-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Power_MatchesExcelDocs()
    {
        // support.microsoft.com POWER: POWER(5,2)=25; POWER(98.6,3.2)=2401077.222;
        // POWER(4,5/4)=5.656854249.
        await Assert.That(Num(Calc("=POWER(5,2)"))).IsEqualTo(25.0);
        await Assert.That(Num(Calc("=POWER(98.6,3.2)"))).IsEqualTo(2401077.222).Within(1e-3);
        await Assert.That(Num(Calc("=POWER(4,5/4)"))).IsEqualTo(5.656854249).Within(1e-9);
    }

    [Test]
    public async Task Power_DomainErrors()
    {
        // Excel: negative base with fractional exponent and 0^0 → #NUM!; 0^negative → #DIV/0!.
        await Assert.That(Calc("=POWER(-2,0.5)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=POWER(0,0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=POWER(0,-2)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Exp_MatchesMath()
    {
        await Assert.That(Num(Calc("=EXP(1)"))).IsEqualTo(Math.E).Within(Tolerance);
        await Assert.That(Num(Calc("=EXP(0)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=EXP(-1)"))).IsEqualTo(Math.Exp(-1)).Within(Tolerance);
    }

    [Test]
    public async Task Exp_OverflowIsNum()
    {
        // Excel numbers cap at ~1E+308; EXP(710) overflows and Excel reports #NUM!.
        await Assert.That(Calc("=EXP(710)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Ln_MatchesMath()
    {
        await Assert.That(Num(Calc("=LN(10)"))).IsEqualTo(Math.Log(10)).Within(Tolerance);
        await Assert.That(Num(Calc("=LN(1)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Ln_NonPositiveIsNum()
    {
        await Assert.That(Calc("=LN(0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=LN(-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Log_MatchesExcelDocs()
    {
        // support.microsoft.com LOG: LOG(10)=1 (default base 10); LOG(8,2)=3;
        // LOG(86,2.7182818)=4.4543473.
        await Assert.That(Num(Calc("=LOG(10)"))).IsEqualTo(1.0).Within(Tolerance);
        await Assert.That(Num(Calc("=LOG(8,2)"))).IsEqualTo(3.0).Within(Tolerance);
        await Assert.That(Num(Calc("=LOG(86,2.7182818)"))).IsEqualTo(4.4543473).Within(1e-7);
    }

    [Test]
    public async Task Log_BaseErrors()
    {
        // Excel (excelfunctions.net/excel-log-function.html, cross-checked with LibreOffice):
        // base = 1 → #DIV/0! (ln(1) = 0); number or base <= 0 → #NUM!.
        await Assert.That(Calc("=LOG(10,1)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=LOG(10,0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=LOG(10,-2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=LOG(0)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=LOG(-5)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Log10_MatchesMath()
    {
        await Assert.That(Num(Calc("=LOG10(100)"))).IsEqualTo(2.0).Within(Tolerance);
        await Assert.That(Num(Calc("=LOG10(7)"))).IsEqualTo(Math.Log10(7)).Within(Tolerance);
        await Assert.That(Calc("=LOG10(0)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task SqrtPi_MatchesExcelDocs()
    {
        // support.microsoft.com SQRTPI: SQRTPI(1)=1.772454; SQRTPI(2)=2.506628; number<0 → #NUM!.
        await Assert.That(Num(Calc("=SQRTPI(1)"))).IsEqualTo(1.772454).Within(1e-6);
        await Assert.That(Num(Calc("=SQRTPI(2)"))).IsEqualTo(2.506628).Within(1e-6);
        await Assert.That(Calc("=SQRTPI(-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task ArgumentErrors_Propagate()
    {
        await Assert.That(Calc("=SQRT(1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=LOG(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=POWER(2,1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }
}
