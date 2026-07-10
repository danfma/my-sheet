using System.IO.Compression;
using System.Text;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Robustness of the streaming merge-join against target files that use IMPLICIT row/cell positions
/// (no @r, allowed by ECMA-376 §18.3.1.4/§18.3.1.73 — a row/cell is the one immediately after the
/// previous element if the reference is omitted). Built as raw ZipArchive parts, like
/// <see cref="StreamingLoadEdgeTests"/>, because ClosedXML always writes explicit @r attributes and so
/// cannot produce this shape. Every case is first proven loadable by <see cref="ExcelFile.Load"/> alone
/// (the loader already tracks implicit positions), then merged and reloaded to prove the merge preserves
/// every existing value at its correct coordinate while our own values land where intended.
/// </summary>
public class ExcelMergeImplicitPositionTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    // Row 1: A1 explicit, B1 implicit (no @r, mid-row — case b), C1 explicit.
    // Row 2: no @r at all (implicit row — case a); A2 explicit column, B2 implicit column.
    // Row 3: B3 explicit, C3 implicit — column 1 (A3) is deliberately free so our workbook can insert a
    //        brand-new cell there, before B3/C3, without disturbing C3's implicit column (case c).
    // Row 6 / Row 7: row 7 has no @r (implicit row 7, following explicit row 6) so our workbook can
    //        insert brand-new rows 4 and 5 before row 6 without disturbing row 7's implicit number
    //        (row-level analogue of case c).
    private const string WorksheetXml = $"""
        <?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="{MainNs}">
        <sheetData>
        <row r="1"><c r="A1"><v>10</v></c><c><v>11</v></c><c r="C1"><v>12</v></c></row>
        <row><c r="A2"><v>20</v></c><c><v>21</v></c></row>
        <row r="3"><c r="B3"><v>30</v></c><c><v>31</v></c></row>
        <row r="6"><c r="A6"><v>60</v></c></row>
        <row><c r="A7"><v>70</v></c></row>
        </sheetData>
        </worksheet>
        """;

    /// <summary>
    /// Builds a minimal single-sheet ("S") .xlsx whose worksheet part is exactly <see cref="WorksheetXml"/>
    /// and returns its path — unlike StreamingLoadEdgeTests' WithRawFixture, the file is left on disk so
    /// the caller can both load AND merge into it.
    /// </summary>
    private static string CreateRawFixture()
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"mysheet-merge-implicit-{Guid.NewGuid():N}.xlsx"
        );

        using (var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create))
        {
            AddEntry(
                zip,
                "[Content_Types].xml",
                $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """
            );
            AddEntry(
                zip,
                "_rels/.rels",
                $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="{PkgRelNs}">
                <Relationship Id="rId1" Type="{RelNs}/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """
            );
            AddEntry(
                zip,
                "xl/workbook.xml",
                $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <workbook xmlns="{MainNs}" xmlns:r="{RelNs}">
                <sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """
            );
            AddEntry(
                zip,
                "xl/_rels/workbook.xml.rels",
                $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="{PkgRelNs}">
                <Relationship Id="rId1" Type="{RelNs}/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """
            );
            AddEntry(zip, "xl/worksheets/sheet1.xml", WorksheetXml);
        }

        return path;

        static void AddEntry(ZipArchive zip, string name, string content)
        {
            using var stream = zip.CreateEntry(name).Open();
            var bytes = Encoding.UTF8.GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }
    }

    private static Workbook BuildWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");

        // Tail cell in row 1 (b): forces MergeRow to walk past B1's implicit position to reach the end
        // of the row.
        sheet["D1"] = ExpressionParser.Parse("=999", sheet);
        sheet["A3"] = ExpressionParser.Parse("=300", sheet); // (c) new cell inserted before B3/C3
        sheet["B3"] = ExpressionParser.Parse("=301", sheet); // (c) replaces the existing explicit B3
        sheet["A4"] = ExpressionParser.Parse("=400", sheet); // (c) new row inserted before row 6
        sheet["A5"] = ExpressionParser.Parse("=500", sheet); // (c) new row inserted before row 6

        return workbook;
    }

    [Test]
    public async Task RawFixture_WithImplicitPositions_LoadsCorrectly()
    {
        // Precondition: the loader already handles implicit rows/cells — this pins down that the fixture
        // itself is spec-compliant and readable before we even touch the merge.
        var path = CreateRawFixture();

        try
        {
            var workbook = ExcelFile.Load(path);

            await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(10.0);
            await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(11.0);
            await Assert.That(workbook.GetCellValue("S", "C1").ToDouble()).IsEqualTo(12.0);
            await Assert.That(workbook.GetCellValue("S", "A2").ToDouble()).IsEqualTo(20.0);
            await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(21.0);
            await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(30.0);
            await Assert.That(workbook.GetCellValue("S", "C3").ToDouble()).IsEqualTo(31.0);
            await Assert.That(workbook.GetCellValue("S", "A6").ToDouble()).IsEqualTo(60.0);
            await Assert.That(workbook.GetCellValue("S", "A7").ToDouble()).IsEqualTo(70.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Merge_TargetWithImplicitRowsAndCells_DoesNotThrow_AndPreservesEveryPosition()
    {
        // Before the fix, MergeSheetData/MergeRow assumed every row/cell carried @r and crashed
        // (NullReferenceException/FormatException) on this exact shape.
        var path = CreateRawFixture();

        try
        {
            BuildWorkbook().MergeIntoExcel(path);

            var merged = ExcelFile.Load(path);

            // Row 1 (case b): the implicit B1 and the explicit A1/C1 around it are untouched, and our
            // brand-new tail cell D1 landed after them.
            await Assert.That(merged.GetCellValue("S", "A1").ToDouble()).IsEqualTo(10.0);
            await Assert.That(merged.GetCellValue("S", "B1").ToDouble()).IsEqualTo(11.0);
            await Assert.That(merged.GetCellValue("S", "C1").ToDouble()).IsEqualTo(12.0);
            await Assert.That(merged.GetCellValue("S", "D1").ToDouble()).IsEqualTo(999.0);

            // Row 2 (case a): the whole row lacked @r and is untouched by us — both its own implicit
            // row number and its implicit B2 column must still resolve correctly.
            await Assert.That(merged.GetCellValue("S", "A2").ToDouble()).IsEqualTo(20.0);
            await Assert.That(merged.GetCellValue("S", "B2").ToDouble()).IsEqualTo(21.0);

            // Row 3 (case c, cell level): our new A3 was inserted before B3/C3, and B3 (explicit) was
            // replaced — neither disturbs the implicit C3, which must still read as column 3.
            await Assert.That(merged.GetCellValue("S", "A3").ToDouble()).IsEqualTo(300.0);
            await Assert.That(merged.GetCellValue("S", "B3").ToDouble()).IsEqualTo(301.0);
            await Assert.That(merged.GetCellValue("S", "C3").ToDouble()).IsEqualTo(31.0);

            // Rows 4-7 (case c, row level): our new rows 4 and 5 were inserted before the existing
            // explicit row 6 — this must not disturb row 7's implicit row number (still "6 + 1").
            await Assert.That(merged.GetCellValue("S", "A4").ToDouble()).IsEqualTo(400.0);
            await Assert.That(merged.GetCellValue("S", "A5").ToDouble()).IsEqualTo(500.0);
            await Assert.That(merged.GetCellValue("S", "A6").ToDouble()).IsEqualTo(60.0);
            await Assert.That(merged.GetCellValue("S", "A7").ToDouble()).IsEqualTo(70.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
