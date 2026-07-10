using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Whole-column (<c>A:A</c>) interop with .xlsx, using ClosedXML (an independent implementation that
/// supports full-column references) as the oracle: our reader parses a <c>&lt;f&gt;</c> holding <c>A:A</c>,
/// and our writer emits <c>A:A</c> so ClosedXML reads it back identically.
/// </summary>
public class WholeColumnInteropTests
{
    [Test]
    public async Task Load_ParsesWholeColumnFormula()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-wholecol-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var fixture = new XLWorkbook())
            {
                var sheet = fixture.AddWorksheet("Data");
                sheet.Cell("A1").Value = 1;
                sheet.Cell("A2").Value = 2;
                sheet.Cell("A3").Value = 3;
                sheet.Cell("C1").FormulaA1 = "SUM(A:A)";
                fixture.SaveAs(path);
            }

            var workbook = ExcelFile.Load(path);

            await Assert.That(workbook.GetCellValue("Data", "C1").ToDouble()).IsEqualTo(6.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Export_WritesWholeColumnFormula_ClosedXmlOracle()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-wholecol-{Guid.NewGuid():N}.xlsx");

        try
        {
            var workbook = new Workbook();
            var data = workbook.Sheets.Add("Data");
            data["A1"] = new NumberValue(4);
            data["A2"] = new NumberValue(6);
            data["C1"] = ExpressionParser.Parse("=SUM(A:A)", data);

            workbook.SaveAsExcel(
                path,
                new ExcelExportOptions { FormulaMode = FormulaMode.Formulas }
            );

            // The oracle reads back the whole-column formula text verbatim.
            using (var oracle = new XLWorkbook(path))
            {
                await Assert
                    .That(oracle.Worksheet("Data").Cell("C1").FormulaA1)
                    .IsEqualTo("SUM(A:A)");
            }

            // And our reader re-evaluates it to the same result.
            var reloaded = ExcelFile.Load(path);
            await Assert.That(reloaded.GetCellValue("Data", "C1").ToDouble()).IsEqualTo(10.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
