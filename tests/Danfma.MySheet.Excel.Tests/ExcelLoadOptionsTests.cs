using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// <see cref="ExcelLoadOptions"/> and its <see cref="ExcelLoadOptions.OnWarning"/> callback. ClosedXML
/// validates defined names on write, so an invalid one is injected directly via the OpenXML SDK after the
/// fixture is saved — the same trick a hand-rolled or third-party .xlsx writer could produce in the wild.
/// </summary>
public class ExcelLoadOptionsTests
{
    private static string WriteFixtureWithInvalidDefinedName()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-warn-{Guid.NewGuid():N}.xlsx");

        using (var fixture = new XLWorkbook())
        {
            fixture.AddWorksheet("Data").Cell("A1").Value = 1;
            fixture.SaveAs(path);
        }

        using (var document = SpreadsheetDocument.Open(path, isEditable: true))
        {
            var workbookPart = document.WorkbookPart!;
            var xlsxWorkbook = workbookPart.Workbook!;

            // ClosedXML always writes an (empty) <definedNames/> element, so append to THAT one rather
            // than appending a second sibling — .DefinedNames only ever returns the first child.
            var definedNames =
                xlsxWorkbook.DefinedNames ?? xlsxWorkbook.AppendChild(new DefinedNames());

            // Unbalanced parenthesis: fails the formula parser, exactly like a hand-edited/foreign-tool
            // .xlsx could carry (ClosedXML itself refuses to write an invalid refersTo).
            definedNames.AppendChild(new DefinedName("SUM(A1:A2") { Name = "BadName" });
            xlsxWorkbook.Save();
        }

        return path;
    }

    [Test]
    public async Task Load_InvalidDefinedName_ReportsWarning_AndWorkbookStillLoads()
    {
        var path = WriteFixtureWithInvalidDefinedName();

        try
        {
            var warnings = new List<ExcelLoadWarning>();
            var options = new ExcelLoadOptions { OnWarning = warnings.Add };

            var workbook = ExcelFile.Load(path, options);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Kind).IsEqualTo(ExcelLoadWarningKind.InvalidDefinedName);
            await Assert.That(warnings[0].Subject).IsEqualTo("BadName");
            await Assert.That(warnings[0].Detail).IsNotEmpty();

            // The rest of the workbook loaded normally; only the bad name was skipped.
            await Assert.That(workbook.DefinedNames.ContainsKey("BadName")).IsFalse();
            await Assert.That(workbook.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(1.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_WithoutOptions_InvalidDefinedName_IsSkippedSilently()
    {
        var path = WriteFixtureWithInvalidDefinedName();

        try
        {
            // No options at all: same call as before ExcelLoadOptions existed. Must not throw, and the
            // load result is identical to the warning-observed case above (the name is just skipped).
            var workbook = ExcelFile.Load(path);

            await Assert.That(workbook.DefinedNames.ContainsKey("BadName")).IsFalse();
            await Assert.That(workbook.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(1.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_Stream_WithOptions_ReportsWarning()
    {
        var path = WriteFixtureWithInvalidDefinedName();

        try
        {
            var warnings = new List<ExcelLoadWarning>();
            var options = new ExcelLoadOptions { OnWarning = warnings.Add };

            using var stream = File.OpenRead(path);
            ExcelFile.Load(stream, options);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Kind).IsEqualTo(ExcelLoadWarningKind.InvalidDefinedName);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_UnparsableDateLiteral_ReportsWarning_AndCellFallsBackToStringValue()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-warn-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var fixture = new XLWorkbook())
            {
                fixture.AddWorksheet("Data").Cell("A1").Value = 1;
                fixture.SaveAs(path);
            }

            // Inject a t="d" cell whose text is not a valid date — the strict-mode ISO-8601 path that
            // DateTime.TryParse rejects. ClosedXML always writes valid dates, so this is done by hand.
            using (var document = SpreadsheetDocument.Open(path, isEditable: true))
            {
                var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                var worksheet = worksheetPart.Worksheet!;
                var sheetData = worksheet.GetFirstChild<SheetData>()!;

                var row = new Row { RowIndex = 2 };
                row.AppendChild(
                    new Cell
                    {
                        CellReference = "B2",
                        DataType = CellValues.Date,
                        CellValue = new CellValue("not-a-date"),
                    }
                );
                sheetData.AppendChild(row);
                worksheet.Save();
            }

            var warnings = new List<ExcelLoadWarning>();
            var options = new ExcelLoadOptions { OnWarning = warnings.Add };

            var workbook = ExcelFile.Load(path, options);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert
                .That(warnings[0].Kind)
                .IsEqualTo(ExcelLoadWarningKind.UnparsableDateLiteral);
            await Assert.That(warnings[0].Subject).IsEqualTo("B2");
            await Assert.That(warnings[0].Detail).IsEqualTo("not-a-date");

            // Same fallback as before this option existed: an unparsable date literal reads as text.
            await Assert.That(workbook.GetCellValue("Data", "B2").ToText()).IsEqualTo("not-a-date");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_ValidWorkbook_NeverInvokesOnWarning()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-warn-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var fixture = new XLWorkbook())
            {
                var data = fixture.AddWorksheet("Data");
                data.Cell("A1").Value = 42;
                data.Range("A1:A1").AddToNamed("GoodName", XLScope.Workbook);
                fixture.SaveAs(path);
            }

            var warningCount = 0;
            var options = new ExcelLoadOptions { OnWarning = _ => warningCount++ };

            var workbook = ExcelFile.Load(path, options);

            await Assert.That(warningCount).IsEqualTo(0);
            await Assert.That(workbook.DefinedNames.ContainsKey("GoodName")).IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }
}
