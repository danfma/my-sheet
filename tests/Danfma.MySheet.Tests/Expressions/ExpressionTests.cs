using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class ExpressionTests
{
    [Test]
    public async Task Sum_ShouldDereferenceCellsToCompute()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Sum(Cell("A1", sheet), Cell("A2", sheet));

        var result = sheet["A3"].Evaluate(workbook).AsObject() as double?;

        await Assert.That(result).IsEqualTo(3);
    }

    [Test]
    public async Task ParseSum_ShouldDereferenceCellsToCompute()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = ExpressionParser.Parse("=SUM(A1,A2)", sheet);

        var result = sheet["A3"].Evaluate(workbook).AsObject() as double?;

        await Assert.That(result).IsEqualTo(3);
    }

    [Test]
    public async Task String_CreatesStringValue()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = String("hello");

        var result = sheet["A1"].Evaluate(workbook).AsObject() as string;

        await Assert.That(result).IsEqualTo("hello");
    }
}
