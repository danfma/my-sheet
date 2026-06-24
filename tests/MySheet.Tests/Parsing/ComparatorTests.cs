using MySheet.Expressions;
using MySheet.Parsing;

namespace MySheet.Tests.Parsing;

public class ComparatorTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Compute(workbook);
    }

    [Test]
    public async Task NumericEquality_StillWorks()
    {
        await Assert.That(Calc("=2=2") as bool?).IsEqualTo(true);
        await Assert.That(Calc("=2<>3") as bool?).IsEqualTo(true);
    }

    [Test]
    public async Task TextEquality_IsCaseInsensitive()
    {
        await Assert.That(Calc("=\"a\"=\"A\"") as bool?).IsEqualTo(true);
        await Assert.That(Calc("=\"a\"=\"b\"") as bool?).IsEqualTo(false);
        await Assert.That(Calc("=\"abc\"<>\"abd\"") as bool?).IsEqualTo(true);
    }

    [Test]
    public async Task MixedTypes_AreNotEqual()
    {
        await Assert.That(Calc("=1=\"1\"") as bool?).IsEqualTo(false);
    }

    [Test]
    public async Task Blank_EqualsZeroAndEmptyString()
    {
        // A1 is blank on an empty sheet.
        await Assert.That(Calc("=A1=0") as bool?).IsEqualTo(true);
        await Assert.That(Calc("=A1=\"\"") as bool?).IsEqualTo(true);
    }

    [Test]
    public async Task TextOrdering_IsCaseInsensitive()
    {
        await Assert.That(Calc("=\"a\"<\"b\"") as bool?).IsEqualTo(true);
        await Assert.That(Calc("=\"B\"<\"a\"") as bool?).IsEqualTo(false); // case-insensitive: B > a
    }

    [Test]
    public async Task CrossTypeOrdering_NumberBeforeTextBeforeBoolean()
    {
        await Assert.That(Calc("=1<\"a\"") as bool?).IsEqualTo(true); // number < text
        await Assert.That(Calc("=\"a\">1") as bool?).IsEqualTo(true);
        await Assert.That(Calc("=\"z\"<TRUE") as bool?).IsEqualTo(true); // text < boolean
    }

    [Test]
    public async Task ErrorPropagates_ThroughEquality()
    {
        await Assert.That(Calc("=(1/0)=1")).IsEqualTo(ErrorValue.DivByZero);
    }
}
