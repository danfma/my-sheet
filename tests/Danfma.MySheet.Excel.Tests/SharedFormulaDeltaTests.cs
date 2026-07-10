using System.IO.Compression;
using System.Text;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Parity freeze for shared-formula expansion: these behaviors were pinned on the ORIGINAL
/// text-shift + reparse pipeline (SharedFormulaShifter) and must keep passing after the expansion
/// moved to token-delta parsing — $-anchored components don't move, function names and sheet
/// qualifiers are never shifted, open-range endpoints stay put, and 2D group deltas shift both axes.
/// </summary>
public class SharedFormulaDeltaTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>Two-sheet raw fixture ("S" under test + "Data") so cross-sheet shifts are pinned.</summary>
    private static async Task WithFixture(string sheetDataXml, Func<Workbook, Task> assert)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-sfd-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create))
            {
                AddEntry(
                    zip,
                    "[Content_Types].xml",
                    """
                    <?xml version="1.0" encoding="UTF-8"?>
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                    <Default Extension="xml" ContentType="application/xml"/>
                    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
                    <sheets>
                    <sheet name="S" sheetId="1" r:id="rId1"/>
                    <sheet name="Data" sheetId="2" r:id="rId2"/>
                    </sheets>
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
                    <Relationship Id="rId2" Type="{RelNs}/worksheet" Target="worksheets/sheet2.xml"/>
                    </Relationships>
                    """
                );
                AddEntry(
                    zip,
                    "xl/worksheets/sheet1.xml",
                    $"""
                    <?xml version="1.0" encoding="UTF-8"?>
                    <worksheet xmlns="{MainNs}">
                    <sheetData>
                    {sheetDataXml}
                    </sheetData>
                    </worksheet>
                    """
                );
                AddEntry(
                    zip,
                    "xl/worksheets/sheet2.xml",
                    $"""
                    <?xml version="1.0" encoding="UTF-8"?>
                    <worksheet xmlns="{MainNs}">
                    <sheetData>
                    <row r="1"><c r="A1"><v>100</v></c></row>
                    <row r="2"><c r="A2"><v>200</v></c></row>
                    <row r="3"><c r="A3"><v>300</v></c></row>
                    </sheetData>
                    </worksheet>
                    """
                );
            }

            await assert(ExcelFile.Load(path));
        }
        finally
        {
            File.Delete(path);
        }

        static void AddEntry(ZipArchive zip, string name, string content)
        {
            using var stream = zip.CreateEntry(name).Open();
            var bytes = Encoding.UTF8.GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }
    }

    [Test]
    public async Task SharedFormula_AbsoluteAnchorsDoNotMove_RelativesDo()
    {
        // Grid: A1=1, A2=10, A3=100. Master B2 = $A$1 + A$1 + $A2 + A2 (every $ combination).
        // Slave B3 (deltaRow=1): $A$1 → $A$1, A$1 → A$1 (col shift only, none here), $A2 → $A3,
        // A2 → A3. Values pin the semantics: master = 1+1+10+10 = 22; slave = 1+1+100+100 = 202.
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>10</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">$A$1+A$1+$A2+A2</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>100</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(22.0);
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(202.0);
            }
        );
    }

    [Test]
    public async Task SharedFormula_TwoDimensionalGroup_ShiftsBothAxes()
    {
        // Master B2 = A1*10 in a B2:C3 group. C3 (deltaCol=1, deltaRow=1) → B2*10; B2's own value is
        // A1*10, so C3 = (A1*10)*10. Mixed anchors across axes: B$1 keeps the row, shifts the column.
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>3</v></c><c r="C1"><v>5</v></c></row>
            <row r="2"><c r="B2"><f t="shared" ref="B2:C3" si="0">A1*10+B$1</f><v>0</v></c><c r="C2"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c><c r="C3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                // B2 = A1*10 + B$1 = 70+3 = 73
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(73.0);
                // C2 (dc=1): B1*10 + C$1 = 30+5 = 35
                await Assert.That(workbook.GetCellValue("S", "C2").ToDouble()).IsEqualTo(35.0);
                // B3 (dr=1): A2*10 + B$1 = 0+3 (A2 blank) = 3
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(3.0);
                // C3 (dc=1, dr=1): B2*10 + C$1 = 730+5 = 735
                await Assert.That(workbook.GetCellValue("S", "C3").ToDouble()).IsEqualTo(735.0);
            }
        );
    }

    [Test]
    public async Task SharedFormula_FunctionNamesAndSheetQualifiers_AreNeverShifted()
    {
        // LOG10 looks like a cell reference but is a function; Data is a sheet qualifier. Only the
        // cell part of Data!A1 shifts (→ Data!A2 on the next row).
        await WithFixture(
            """
            <row r="2"><c r="B2"><f t="shared" ref="B2:B3" si="0">LOG10(100)+Data!A1</f><v>0</v></c></row>
            <row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(102.0);
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(202.0);
            }
        );
    }

    [Test]
    public async Task SharedFormula_OpenRangeEndpoints_StayPut()
    {
        // A:A endpoints are letters-only tokens — never shifted (only letters+digits tokens are).
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">SUM(A:A)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>4</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(7.0);
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(7.0);
            }
        );
    }

    [Test]
    public async Task SharedFormula_StringLiterals_AreCopiedUntouched()
    {
        // "A1" inside a string literal must not shift; the concatenated text pins it.
        await WithFixture(
            """
            <row r="2"><c r="B2"><f t="shared" ref="B2:B3" si="0">"A1 says "&amp;A1</f><v>0</v></c></row>
            <row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c><c r="A3" t="str"><v>x</v></c></row>
            <row r="4"><c r="A4" t="str"><v>y</v></c></row>
            """,
            async workbook =>
            {
                // B3 (dr=1): "A1 says " & A2 — A2 is blank → empty coercion; B2 refs A1 (blank too).
                // The literal prefix is the pinned part.
                await Assert.That(workbook.GetCellValue("S", "B2").ToText()).IsEqualTo("A1 says ");
                await Assert.That(workbook.GetCellValue("S", "B3").ToText()).IsEqualTo("A1 says ");
            }
        );
    }

    [Test]
    public async Task SharedFormula_RangeReference_ShiftsBothEndpoints()
    {
        // A1:A2 in the master (B3) becomes A2:A3 in the slave (B4).
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>10</v></c></row>
            <row r="3"><c r="A3"><v>100</v></c><c r="B3"><f t="shared" ref="B3:B4" si="0">SUM(A1:A2)</f><v>0</v></c></row>
            <row r="4"><c r="B4"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(11.0);
                await Assert.That(workbook.GetCellValue("S", "B4").ToDouble()).IsEqualTo(110.0);
            }
        );
    }
}
