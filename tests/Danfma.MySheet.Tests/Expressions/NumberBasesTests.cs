using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// BASE/DECIMAL/ROMAN/ARABIC. Golden values from the Microsoft support pages (fetched 2026-07-01).
/// ROMAN is implemented in the classic form only (form 0/TRUE/omitted); the concise forms 1-4 and
/// FALSE are a documented limitation and return #VALUE!.
/// </summary>
public class NumberBasesTests
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
    public async Task Base_MatchesExcelDocs()
    {
        // support.microsoft.com BASE: BASE(7,2)="111"; BASE(100,16)="64"; BASE(15,2,10)="0000001111".
        await Assert.That(Calc("=BASE(7,2)")).IsEqualTo("111");
        await Assert.That(Calc("=BASE(100,16)")).IsEqualTo("64");
        await Assert.That(Calc("=BASE(15,2,10)")).IsEqualTo("0000001111");
        await Assert.That(Calc("=BASE(0,2)")).IsEqualTo("0");
        await Assert.That(Calc("=BASE(35,36)")).IsEqualTo("Z");
    }

    [Test]
    public async Task Base_ConstraintViolationsAreNum()
    {
        // support.microsoft.com BASE: number in [0, 2^53), radix in [2, 36], min_length in [0, 255];
        // outside the constraints → #NUM!.
        await Assert.That(Calc("=BASE(-1,2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=BASE(9007199254740992,2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=BASE(7,1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=BASE(7,37)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=BASE(7,2,-1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=BASE(7,2,256)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Decimal_MatchesExcelDocs()
    {
        // support.microsoft.com DECIMAL: DECIMAL("FF",16)=255; DECIMAL(111,2)=7;
        // DECIMAL("zap",36)=45745 — case-insensitive.
        await Assert.That(Num(Calc("=DECIMAL(\"FF\",16)"))).IsEqualTo(255.0);
        await Assert.That(Num(Calc("=DECIMAL(\"ff\",16)"))).IsEqualTo(255.0);
        await Assert.That(Num(Calc("=DECIMAL(111,2)"))).IsEqualTo(7.0);
        await Assert.That(Num(Calc("=DECIMAL(\"zap\",36)"))).IsEqualTo(45745.0);
    }

    [Test]
    public async Task Decimal_InvalidInputErrors()
    {
        // Digits invalid for the radix, or radix outside [2, 36] → #NUM!.
        await Assert.That(Calc("=DECIMAL(\"FF\",2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=DECIMAL(\"10\",1)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=DECIMAL(\"10\",37)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=DECIMAL(\"1.5\",10)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Roman_ClassicForm()
    {
        // support.microsoft.com ROMAN: ROMAN(499,0)="CDXCIX" (classic); TRUE also means classic.
        // ROMAN(1912)="MCMXII" cross-checks the documented ARABIC("mcmxii")=1912.
        await Assert.That(Calc("=ROMAN(499)")).IsEqualTo("CDXCIX");
        await Assert.That(Calc("=ROMAN(499,0)")).IsEqualTo("CDXCIX");
        await Assert.That(Calc("=ROMAN(499,TRUE)")).IsEqualTo("CDXCIX");
        await Assert.That(Calc("=ROMAN(1912)")).IsEqualTo("MCMXII");
        await Assert.That(Calc("=ROMAN(3999)")).IsEqualTo("MMMCMXCIX");
        // ROMAN(0) is the empty string (ODF OpenFormula; Excel behaviour).
        await Assert.That(Calc("=ROMAN(0)")).IsEqualTo("");
    }

    [Test]
    public async Task Roman_OutOfRangeIsValueError()
    {
        // support.microsoft.com ROMAN: negative or greater than 3999 → #VALUE!.
        await Assert.That(Calc("=ROMAN(-1)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ROMAN(4000)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Roman_ConciseFormsAreAnUnsupportedLimitation()
    {
        // Deliberate limitation: Excel's concise forms (1-4/FALSE, e.g. ROMAN(499,1)="LDVLIV")
        // are not implemented because their algorithm is not verifiable against a full oracle;
        // MySheet returns #VALUE! instead of a wrong numeral.
        await Assert.That(Calc("=ROMAN(499,1)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ROMAN(499,4)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ROMAN(499,FALSE)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Arabic_MatchesExcelDocs()
    {
        // support.microsoft.com ARABIC: ARABIC("LVII")=57; "mcmxii"=1912 (case ignored);
        // ""=0; leading '-' negates ("-MMXI"=-2011); leading/trailing spaces ignored.
        await Assert.That(Num(Calc("=ARABIC(\"LVII\")"))).IsEqualTo(57.0);
        await Assert.That(Num(Calc("=ARABIC(\"mcmxii\")"))).IsEqualTo(1912.0);
        await Assert.That(Num(Calc("=ARABIC(\"\")"))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=ARABIC(\"-MMXI\")"))).IsEqualTo(-2011.0);
        await Assert.That(Num(Calc("=ARABIC(\" LVII \")"))).IsEqualTo(57.0);
    }

    [Test]
    public async Task Arabic_InvalidInputIsValueError()
    {
        // support.microsoft.com ARABIC: numbers and text that is not a valid Roman numeral → #VALUE!.
        await Assert.That(Calc("=ARABIC(\"ABC\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ARABIC(123)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=ARABIC(\"-\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task ArgumentErrors_Propagate()
    {
        await Assert.That(Calc("=BASE(1/0,2)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=ROMAN(1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }
}
