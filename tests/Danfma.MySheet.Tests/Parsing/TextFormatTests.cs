using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class TextFormatTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task Text_NumericFormats()
    {
        await Assert.That(Calc("=TEXT(1234.5,\"#,##0.00\")") as string).IsEqualTo("1,234.50");
        await Assert.That(Calc("=TEXT(0.5,\"0%\")") as string).IsEqualTo("50%");
    }

    [Test]
    public async Task Text_DateFormats()
    {
        // 44197 is the Excel serial for 2021-01-01.
        await Assert.That(Calc("=TEXT(44197,\"yyyy-mm-dd\")") as string).IsEqualTo("2021-01-01");
        await Assert.That(Calc("=TEXT(44197,\"dd/mm/yyyy\")") as string).IsEqualTo("01/01/2021");
        await Assert.That(Calc("=TEXT(44197,\"mmm yyyy\")") as string).IsEqualTo("Jan 2021");
    }

    [Test]
    public async Task Text_TimeFormats()
    {
        // 0.5 of a day is noon; mm here is minutes (it follows hh).
        await Assert.That(Calc("=TEXT(0.5,\"hh:mm:ss\")") as string).IsEqualTo("12:00:00");
    }
}
