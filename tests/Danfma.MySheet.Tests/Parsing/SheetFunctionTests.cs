using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class SheetFunctionTests
{
    [Test]
    public async Task Sheet_ByNameAndCurrentSheet()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1"); // index 0
        workbook.Sheets.Add("Sheet2"); // index 1

        await Assert
            .That(
                ExpressionParser.Parse("=SHEET(\"Sheet2\")", sheet1).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(2.0);

        // SHEET() with no argument uses the cell's own sheet (reached through a reference).
        sheet1["A1"] = ExpressionParser.Parse("=SHEET()", sheet1);

        await Assert
            .That(ExpressionParser.Parse("=A1", sheet1).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(1.0);
    }
}
