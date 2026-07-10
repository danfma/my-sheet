using ClosedXML.Excel;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Interop guard for the real-world scenario from the bug report: an .xlsx whose formulas carry dangling
/// cross-sheet references (a sheet that is not present in the file). Loading and evaluating such a workbook
/// must resolve each dangling cell to <c>#REF!</c> — never throw <see cref="KeyNotFoundException"/> and abort
/// the batch.
/// </summary>
public class MissingSheetInteropTests
{
    [Test]
    public async Task Load_FormulaReferencingMissingSheet_EvaluatesToRef()
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"mysheet-missingsheet-{Guid.NewGuid():N}.xlsx"
        );

        try
        {
            using (var fixture = new XLWorkbook())
            {
                var sheet = fixture.AddWorksheet("Data");
                sheet.Cell("A1").Value = 10;
                // Formulas referencing a sheet ("Ghost") that is NOT in the file.
                sheet.Cell("B1").FormulaA1 = "Ghost!A1";
                sheet.Cell("B2").FormulaA1 = "SUM(Ghost!A:A)";
                sheet.Cell("B3").FormulaA1 = "UPPER(Ghost!E9)";
                fixture.SaveAs(path);
            }

            var workbook = ExcelFile.Load(path);

            // The whole batch resolves without throwing; each dangling cell is #REF!.
            foreach (var id in new[] { "B1", "B2", "B3" })
            {
                var value = workbook.GetCellValue("Data", id);

                await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
                value.TryGetError(out var error);
                await Assert.That(error).IsEqualTo(Error.Ref);
            }

            // A well-formed cell in the same workbook still evaluates normally.
            await Assert.That(workbook.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(10.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
