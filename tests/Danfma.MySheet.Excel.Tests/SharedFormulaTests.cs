using Danfma.MySheet.Excel;
using Danfma.MySheet.Expressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using XlsxRow = DocumentFormat.OpenXml.Spreadsheet.Row;
using XlsxWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Real Excel files store dragged formulas as SHARED formulas: the master cell carries the text
/// (<c>&lt;f t="shared" si ref&gt;A1*2&lt;/f&gt;</c>) and the other cells only reference the group
/// (<c>&lt;f t="shared" si/&gt;</c>). ClosedXML never writes them, so the fixture is raw OpenXML.
/// </summary>
public class SharedFormulaTests
{
    private static Cell Number(string id, double value) =>
        new() { CellReference = id, CellValue = new CellValue(value.ToString()) };

    private static Cell Master(string id, string formula, uint index, string range, string cached) =>
        new()
        {
            CellReference = id,
            CellFormula = new CellFormula(formula)
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = index,
                Reference = range,
            },
            CellValue = new CellValue(cached),
        };

    private static Cell Slave(string id, uint index, string cached) =>
        new()
        {
            CellReference = id,
            CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = index },
            CellValue = new CellValue(cached),
        };

    /// <summary>A1..A3 = 1..3; B = shared "A1*2" (relative); C = shared "$A$1+A1" (absolute + relative);
    /// D = shared "&quot;A1&quot;&amp;A1" (a ref-lookalike inside a string literal).</summary>
    private static string CreateFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-shared-{Guid.NewGuid():N}.xlsx");

        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new XlsxWorkbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetData(
                new XlsxRow(
                    Number("A1", 1),
                    Master("B1", "A1*2", 0, "B1:B3", "2"),
                    Master("C1", "$A$1+A1", 1, "C1:C2", "2"),
                    Master("D1", "\"A1\"&A1", 2, "D1:D2", "A11")
                ) { RowIndex = 1 },
                new XlsxRow(
                    Number("A2", 2),
                    Slave("B2", 0, "4"),
                    Slave("C2", 1, "3"),
                    Slave("D2", 2, "A12")
                ) { RowIndex = 2 },
                new XlsxRow(Number("A3", 3), Slave("B3", 0, "6")) { RowIndex = 3 }
            )
        );

        workbookPart.Workbook.AppendChild(new Sheets()).AppendChild(
            new XlsxSheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" }
        );

        return path;
    }

    private static async Task WithFixture(Func<Workbook, Task> assert)
    {
        var path = CreateFixture();

        try
        {
            await assert(ExcelFile.Load(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task SlaveCells_BecomeRealFormulas_WithRelativeRefsShifted()
    {
        await WithFixture(async workbook =>
        {
            // Slaves are FORMULAS (shifted from the master), not frozen literals.
            await Assert.That(workbook["Data"]["B2"]).IsTypeOf<BinaryOperation>();
            await Assert.That(workbook["Data"]["B3"]).IsTypeOf<BinaryOperation>();
            await Assert.That(workbook.GetCellValue("Data", "B2").ToDouble()).IsEqualTo(4.0);
            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(6.0);
        });
    }

    [Test]
    public async Task SlaveFormulas_ReactToInputChanges()
    {
        await WithFixture(async workbook =>
        {
            // The whole point of the expansion: change an input and the slave recomputes —
            // a frozen cached literal would keep answering 4.
            workbook["Data"]["A2"] = new NumberValue(10);
            workbook.InvalidateCache();

            await Assert.That(workbook.GetCellValue("Data", "B2").ToDouble()).IsEqualTo(20.0);
        });
    }

    [Test]
    public async Task AbsoluteRefs_DoNotShift()
    {
        await WithFixture(async workbook =>
        {
            // C2 = $A$1+A2 → 1+2. If the absolute part had shifted (A2+A2=4) this fails.
            await Assert.That(workbook.GetCellValue("Data", "C2").ToDouble()).IsEqualTo(3.0);
        });
    }

    [Test]
    public async Task RefLookalikes_InsideStrings_DoNotShift()
    {
        await WithFixture(async workbook =>
        {
            // D2 = "A1"&A2 → the quoted "A1" is text and must not become "A2".
            await Assert.That(workbook.GetCellValue("Data", "D2").ToText()).IsEqualTo("A12");
        });
    }
}
