using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ComparatorTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task NumericEquality_StillWorks()
    {
        await Assert.That(Calc("=2=2") as bool?).IsTrue();
        await Assert.That(Calc("=2<>3") as bool?).IsTrue();
    }

    [Test]
    public async Task TextEquality_IsCaseInsensitive()
    {
        await Assert.That(Calc("=\"a\"=\"A\"") as bool?).IsTrue();
        await Assert.That(Calc("=\"a\"=\"b\"") as bool?).IsFalse();
        await Assert.That(Calc("=\"abc\"<>\"abd\"") as bool?).IsTrue();
    }

    [Test]
    public async Task MixedTypes_AreNotEqual()
    {
        await Assert.That(Calc("=1=\"1\"") as bool?).IsFalse();
    }

    [Test]
    public async Task Blank_EqualsZeroAndEmptyString()
    {
        // A1 is blank on an empty sheet.
        await Assert.That(Calc("=A1=0") as bool?).IsTrue();
        await Assert.That(Calc("=A1=\"\"") as bool?).IsTrue();
    }

    [Test]
    public async Task TextOrdering_IsCaseInsensitive()
    {
        await Assert.That(Calc("=\"a\"<\"b\"") as bool?).IsTrue();
        await Assert.That(Calc("=\"B\"<\"a\"") as bool?).IsFalse(); // case-insensitive: B > a
    }

    [Test]
    public async Task CrossTypeOrdering_NumberBeforeTextBeforeBoolean()
    {
        await Assert.That(Calc("=1<\"a\"") as bool?).IsTrue(); // number < text
        await Assert.That(Calc("=\"a\">1") as bool?).IsTrue();
        await Assert.That(Calc("=\"z\"<TRUE") as bool?).IsTrue(); // text < boolean
    }

    [Test]
    public async Task ErrorPropagates_ThroughEquality()
    {
        await Assert.That(Calc("=(1/0)=1")).IsEqualTo(ErrorValue.DivByZero);
    }
}
