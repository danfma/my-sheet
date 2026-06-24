using MySheet.Parsing;

namespace MySheet.Tests.Parsing;

public class TextFormatTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Compute(workbook);
    }

    [Test]
    public async Task Text_NumericFormats()
    {
        await Assert.That(Calc("=TEXT(1234.5,\"#,##0.00\")") as string).IsEqualTo("1,234.50");
        await Assert.That(Calc("=TEXT(0.5,\"0%\")") as string).IsEqualTo("50%");
    }
}
