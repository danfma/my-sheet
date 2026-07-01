using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class MathInfoFunctionTests
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

    [Test]
    public async Task Int_FloorsTowardNegativeInfinity()
    {
        await Assert.That(Calc("=INT(2.9)") as double?).IsEqualTo(2.0);
        await Assert.That(Calc("=INT(-2.1)") as double?).IsEqualTo(-3.0);
    }

    [Test]
    public async Task Round_HalfAwayFromZero()
    {
        await Assert.That(Calc("=ROUND(3.14159,2)") as double?).IsEqualTo(3.14);
        await Assert.That(Calc("=ROUND(2.5,0)") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=ROUND(-2.5,0)") as double?).IsEqualTo(-3.0);
    }

    [Test]
    public async Task RoundUp_AwayFromZero()
    {
        await Assert.That(Calc("=ROUNDUP(2.001,2)") as double?).IsEqualTo(2.01);
        await Assert.That(Calc("=ROUNDUP(-2.001,2)") as double?).IsEqualTo(-2.01);
    }

    [Test]
    public async Task Abs_Magnitude()
    {
        await Assert.That(Calc("=ABS(-3)") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=ABS(3)") as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task IsNumber_ChecksType()
    {
        await Assert.That(Calc("=ISNUMBER(1)") as bool?).IsTrue();
        await Assert.That(Calc("=ISNUMBER(\"x\")") as bool?).IsFalse();
        await Assert.That(Calc("=ISNUMBER(1/0)") as bool?).IsFalse();
    }

    [Test]
    public async Task IsBlank_DetectsEmptyCell()
    {
        await Assert.That(Calc("=ISBLANK(A1)") as bool?).IsTrue();
        await Assert.That(Calc("=ISBLANK(0)") as bool?).IsFalse();
    }

    [Test]
    public async Task IfNa_OnlyCatchesNotAvailable()
    {
        await Assert.That(Calc("=IFNA(5,-1)") as double?).IsEqualTo(5.0);
        // #DIV/0! is not #N/A, so IFNA passes it through.
        await Assert.That(Calc("=IFNA(1/0,-1)")).IsEqualTo(ErrorValue.DivByZero);
    }
}
