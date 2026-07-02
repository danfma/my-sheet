using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// FACT/FACTDOUBLE/COMBIN/COMBINA/GCD/LCM. Golden values from the Microsoft support pages
/// (fetched 2026-07-01).
/// </summary>
public class CombinatoricsTests
{
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
    public async Task Fact_MatchesExcelDocs()
    {
        // support.microsoft.com FACT: FACT(5)=120; FACT(1.9)=1 (truncated); FACT(0)=1; FACT(1)=1;
        // FACT(-1)=#NUM!.
        await Assert.That(Num(Calc("=FACT(5)"))).IsEqualTo(120.0);
        await Assert.That(Num(Calc("=FACT(1.9)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=FACT(0)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=FACT(1)"))).IsEqualTo(1.0);
        await Assert.That(Calc("=FACT(-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Fact_OverflowIsNum()
    {
        // 171! exceeds the double range (~1E+308) → #NUM!, like Excel's numeric overflow.
        await Assert.That(Calc("=FACT(171)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Num(Calc("=FACT(170)"))).IsEqualTo(7.257415615307994e306).Within(1e293);
    }

    [Test]
    public async Task FactDouble_MatchesExcelDocs()
    {
        // support.microsoft.com FACTDOUBLE: FACTDOUBLE(6)=48 (6·4·2); FACTDOUBLE(7)=105 (7·5·3);
        // negative → #NUM!; non-integers truncated.
        await Assert.That(Num(Calc("=FACTDOUBLE(6)"))).IsEqualTo(48.0);
        await Assert.That(Num(Calc("=FACTDOUBLE(7)"))).IsEqualTo(105.0);
        await Assert.That(Num(Calc("=FACTDOUBLE(0)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=FACTDOUBLE(7.9)"))).IsEqualTo(105.0);
        await Assert.That(Calc("=FACTDOUBLE(-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Combin_MatchesExcelDocs()
    {
        // support.microsoft.com COMBIN: COMBIN(8,2)=28; arguments truncated to integers;
        // n<0, k<0 or n<k → #NUM!.
        await Assert.That(Num(Calc("=COMBIN(8,2)"))).IsEqualTo(28.0);
        await Assert.That(Num(Calc("=COMBIN(8.9,2.9)"))).IsEqualTo(28.0);
        await Assert.That(Num(Calc("=COMBIN(4,4)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=COMBIN(4,0)"))).IsEqualTo(1.0);
        await Assert.That(Calc("=COMBIN(2,3)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=COMBIN(-1,1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=COMBIN(4,-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task CombinA_MatchesExcelDocs()
    {
        // support.microsoft.com COMBINA: COMBINA(4,3)=20; COMBINA(10,3)=220
        // (COMBINA(n,k) = COMBIN(n+k-1,k)); negative arguments → #NUM!.
        await Assert.That(Num(Calc("=COMBINA(4,3)"))).IsEqualTo(20.0);
        await Assert.That(Num(Calc("=COMBINA(10,3)"))).IsEqualTo(220.0);
        await Assert.That(Num(Calc("=COMBINA(4,0)"))).IsEqualTo(1.0);
        await Assert.That(Calc("=COMBINA(-1,2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=COMBINA(4,-1)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Gcd_MatchesExcelDocs()
    {
        // support.microsoft.com GCD: GCD(5,2)=1; GCD(24,36)=12; GCD(7,1)=1; GCD(5,0)=5;
        // non-integers truncated; negative → #NUM!; parameter >= 2^53 → #NUM!.
        await Assert.That(Num(Calc("=GCD(5,2)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=GCD(24,36)"))).IsEqualTo(12.0);
        await Assert.That(Num(Calc("=GCD(7,1)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=GCD(5,0)"))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=GCD(24.9,36.9)"))).IsEqualTo(12.0);
        await Assert.That(Calc("=GCD(-1,5)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=GCD(9007199254740992,2)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Gcd_IsRangeAware()
    {
        await Assert.That(
            Num(Calc("=GCD(A1:A2)", ("A1", new NumberValue(24)), ("A2", new NumberValue(36))))
        ).IsEqualTo(12.0);
    }

    [Test]
    public async Task Lcm_MatchesExcelDocs()
    {
        // support.microsoft.com LCM: LCM(5,2)=10; LCM(24,36)=72; negative → #NUM!;
        // LCM(a,b) >= 2^53 → #NUM!. A zero argument yields 0 (ODF OpenFormula LCM).
        await Assert.That(Num(Calc("=LCM(5,2)"))).IsEqualTo(10.0);
        await Assert.That(Num(Calc("=LCM(24,36)"))).IsEqualTo(72.0);
        await Assert.That(Num(Calc("=LCM(5,0)"))).IsEqualTo(0.0);
        await Assert.That(Calc("=LCM(-1,5)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=LCM(4503599627370496,4503599627370497)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task ArgumentErrors_Propagate()
    {
        await Assert.That(Calc("=FACT(1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=COMBIN(\"abc\",2)")).IsEqualTo(ErrorValue.NotValue);
    }
}
