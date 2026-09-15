using ClosedXML.Excel;
using Danfma.MySheet.Excel;

namespace Danfma.MySheet.Excel.Tests;

public class ArrayConstantInteropTests
{
    [Test]
    public async Task ArrayConstant_LoadsEvaluatesAndExportsAsFormula()
    {
        var source = Path.Combine(
            Path.GetTempPath(),
            $"mysheet-array-source-{Guid.NewGuid():N}.xlsx"
        );
        var exported = Path.Combine(
            Path.GetTempPath(),
            $"mysheet-array-export-{Guid.NewGuid():N}.xlsx"
        );

        try
        {
            using (var fixture = new XLWorkbook())
            {
                fixture.AddWorksheet("Data").Cell("A1").FormulaA1 = "SUM({1,2,3})";
                fixture.SaveAs(source);
            }

            var workbook = ExcelFile.Load(source);
            await Assert.That(workbook.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(6d);

            workbook.SaveAsExcel(
                exported,
                new ExcelExportOptions { FormulaMode = FormulaMode.Formulas }
            );
            using var oracle = new XLWorkbook(exported);
            await Assert
                .That(oracle.Worksheet("Data").Cell("A1").FormulaA1)
                .IsEqualTo("SUM({1,2,3})");

            var reloaded = ExcelFile.Load(exported);
            await Assert.That(reloaded.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(6d);
        }
        finally
        {
            File.Delete(source);
            File.Delete(exported);
        }
    }
}
