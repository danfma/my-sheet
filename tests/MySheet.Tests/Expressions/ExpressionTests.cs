using MySheet.Parsing;
using static MySheet.Expressions.Expression;

namespace MySheet.Tests.Expressions;

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

        var result = sheet["A3"].Compute(workbook) as double?;

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

        var result = sheet["A3"].Compute(workbook) as double?;

        await Assert.That(result).IsEqualTo(3);
    }
}
