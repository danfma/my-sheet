using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class TextFunctionTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Compute(workbook);
    }

    [Test]
    public async Task Upper_And_Lower()
    {
        await Assert.That(Calc("=UPPER(\"ab\")") as string).IsEqualTo("AB");
        await Assert.That(Calc("=LOWER(\"AB\")") as string).IsEqualTo("ab");
    }

    [Test]
    public async Task Trim_CollapsesInternalSpaces()
    {
        await Assert.That(Calc("=TRIM(\" a  b \")") as string).IsEqualTo("a b");
    }

    [Test]
    public async Task Len_CountsCharacters()
    {
        await Assert.That(Calc("=LEN(\"abc\")") as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Left_WithAndWithoutCount()
    {
        await Assert.That(Calc("=LEFT(\"abcd\",2)") as string).IsEqualTo("ab");
        await Assert.That(Calc("=LEFT(\"abcd\")") as string).IsEqualTo("a");
        await Assert.That(Calc("=LEFT(\"ab\",5)") as string).IsEqualTo("ab");
    }

    [Test]
    public async Task Mid_Substring()
    {
        await Assert.That(Calc("=MID(\"abcde\",2,3)") as string).IsEqualTo("bcd");
    }

    [Test]
    public async Task Value_ParsesNumber()
    {
        await Assert.That(Calc("=VALUE(\"12\")") as double?).IsEqualTo(12.0);
        await Assert.That(Calc("=VALUE(\"1.5\")") as double?).IsEqualTo(1.5);
    }

    [Test]
    public async Task Concat_And_Concatenate()
    {
        await Assert.That(Calc("=CONCAT(\"a\",\"b\",\"c\")") as string).IsEqualTo("abc");
        await Assert.That(Calc("=CONCATENATE(\"a\",1)") as string).IsEqualTo("a1");
    }

    [Test]
    public async Task TextJoin_RespectsIgnoreEmpty()
    {
        await Assert
            .That(Calc("=TEXTJOIN(\"-\",TRUE,\"a\",\"\",\"b\")") as string)
            .IsEqualTo("a-b");
        await Assert
            .That(Calc("=TEXTJOIN(\"-\",FALSE,\"a\",\"\",\"b\")") as string)
            .IsEqualTo("a--b");
    }

    [Test]
    public async Task TextFunctions_PropagateErrors()
    {
        await Assert.That(Calc("=UPPER(1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }
}
