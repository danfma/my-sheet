using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Edge cases of the streaming .xlsx loader that ClosedXML cannot produce (implicit cell positions,
/// namespace prefixes, out-of-spec shared-formula ordering, raw cached values), written as raw parts
/// via ZipArchive, plus whitespace/rich-text fidelity through a real independent producer.
/// </summary>
public class StreamingLoadEdgeTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// Builds a minimal single-sheet ("S") .xlsx whose worksheet part is EXACTLY the given XML
    /// document, loads it with our reader, runs the assertions and deletes the file. Tests control
    /// prefixes, attribute presence and element ordering — everything a DOM writer would normalize.
    /// </summary>
    private static async Task WithRawFixture(
        string worksheetXml,
        Func<Workbook, Task> assert,
        string? sharedStringsXml = null
    )
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-raw-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create))
            {
                var sharedStringsOverride = sharedStringsXml is not null
                    ? """<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"""
                    : string.Empty;
                var sharedStringsRel = sharedStringsXml is not null
                    ? $"""<Relationship Id="rId2" Type="{RelNs}/sharedStrings" Target="sharedStrings.xml"/>"""
                    : string.Empty;

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
                    {sharedStringsOverride}
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
                    {sharedStringsRel}
                    </Relationships>
                    """
                );
                AddEntry(zip, "xl/worksheets/sheet1.xml", worksheetXml);

                if (sharedStringsXml is not null)
                {
                    AddEntry(zip, "xl/sharedStrings.xml", sharedStringsXml);
                }
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
    public async Task Load_CellsWithoutReference_UseImplicitPositions()
    {
        // Rows and cells without @r are positioned implicitly (next row / next column) — allowed by
        // the spec, produced by minimal writers.
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row><c><v>1</v></c><c><v>2</v></c></row>
            <row><c><v>3</v></c><c r="D2"><v>4</v></c><c><v>5</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(1.0);
                await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(2.0);
                await Assert.That(workbook.GetCellValue("S", "A2").ToDouble()).IsEqualTo(3.0);
                await Assert.That(workbook.GetCellValue("S", "D2").ToDouble()).IsEqualTo(4.0);
                // After an explicit D2, the next implicit cell is E2.
                await Assert.That(workbook.GetCellValue("S", "E2").ToDouble()).IsEqualTo(5.0);
            }
        );
    }

    [Test]
    public async Task Load_NamespacePrefixedWorksheet_IsRead()
    {
        // Some producers prefix every element (x:worksheet); matching is by local name.
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <x:worksheet xmlns:x="{MainNs}">
            <x:sheetData>
            <x:row r="1"><x:c r="A1"><x:v>42</x:v></x:c><x:c r="B1"><x:f>A1*2</x:f><x:v>0</x:v></x:c></x:row>
            </x:sheetData>
            </x:worksheet>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(42.0);
                await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(84.0);
            }
        );
    }

    [Test]
    public async Task Load_SharedFormulaSlaveBeforeMaster_ResolvesAgainstTheMaster()
    {
        // Out-of-spec producer: the slave (si=0, no text) appears before its master. The deferred
        // resolution still expands it from the master, shifting relative references (A2's B2*2 → B1*2).
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1"><f t="shared" si="0"/><v>99</v></c><c r="B1"><v>10</v></c></row>
            <row r="2"><c r="A2"><f t="shared" ref="A1:A2" si="0">B2*2</f><v>0</v></c><c r="B2"><v>7</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(20.0);
                await Assert.That(workbook.GetCellValue("S", "A2").ToDouble()).IsEqualTo(14.0);
            }
        );
    }

    [Test]
    public async Task Load_SharedFormulaSlaveWithoutMaster_FallsBackToCachedLiteral()
    {
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1"><f t="shared" si="99"/><v>7</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(7.0);
            }
        );
    }

    [Test]
    public async Task Load_EmptyNonSharedFormula_FallsBackToCachedLiteral()
    {
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1"><f/><v>7</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(7.0);
            }
        );
    }

    [Test]
    public async Task Load_InlineString_PreservesWhitespace()
    {
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> inline  text </t></is></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert
                    .That(workbook.GetCellValue("S", "A1").ToText())
                    .IsEqualTo(" inline  text ");
            }
        );
    }

    [Test]
    public async Task Load_SharedString_RichTextAndWhitespace_AreFlattenedFaithfully()
    {
        // si 0: plain padded text; si 1: rich-text runs (bold + preserved-space run) flatten by
        // concatenating every <t>, exactly like the DOM's InnerText.
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert
                    .That(workbook.GetCellValue("S", "A1").ToText())
                    .IsEqualTo("  padded  ");
                await Assert
                    .That(workbook.GetCellValue("S", "B1").ToText())
                    .IsEqualTo("synthetic rich text ");
            },
            sharedStringsXml: $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <sst xmlns="{MainNs}" count="2" uniqueCount="2">
            <si><t xml:space="preserve">  padded  </t></si>
            <si><r><rPr><b/></rPr><t>synthetic</t></r><r><t xml:space="preserve"> rich text </t></r></si>
            </sst>
            """
        );
    }

    [Test]
    public async Task Load_CachedStringResult_WithoutFormula_ReadsAsText()
    {
        // t="str" is a formula's cached string result; without an <f> it degrades to plain text.
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1" t="str"><v>cached result</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert
                    .That(workbook.GetCellValue("S", "A1").ToText())
                    .IsEqualTo("cached result");
            }
        );
    }

    [Test]
    public async Task Load_IsoDate_BecomesSerialNumber()
    {
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <sheetData>
            <row r="1"><c r="A1" t="d"><v>2026-07-10</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert
                    .That(workbook.GetCellValue("S", "A1").ToDouble())
                    .IsEqualTo(new DateTime(2026, 7, 10).ToOADate());
            }
        );
    }

    [Test]
    public async Task Load_HugeDimensionOnSparseSheet_IsHarmless()
    {
        // The dimension bbox is only a (capped) presize hint: a pathological bbox over a 1-cell
        // sheet must not affect correctness (nor balloon memory — the cap bounds the reservation).
        await WithRawFixture(
            $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="{MainNs}">
            <dimension ref="A1:XFD1048576"/>
            <sheetData>
            <row r="1"><c r="A1"><v>7</v></c></row>
            </sheetData>
            </worksheet>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "A1").ToDouble()).IsEqualTo(7.0);
                await Assert.That(workbook["S"].Count).IsEqualTo(1);
            }
        );
    }

    [Test]
    public async Task Load_Export_Load_RoundTripsValues()
    {
        // load(export(load(x))): computed values survive a full round trip through our own writer.
        var source = Path.Combine(Path.GetTempPath(), $"mysheet-rt-src-{Guid.NewGuid():N}.xlsx");
        var exported = Path.Combine(Path.GetTempPath(), $"mysheet-rt-out-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var fixture = new XLWorkbook())
            {
                var sheet = fixture.AddWorksheet("Data");
                sheet.Cell("A1").Value = 21;
                sheet.Cell("A2").FormulaA1 = "A1*2";
                sheet.Cell("A3").Value = "  padded  ";
                fixture.SaveAs(source);
            }

            var first = ExcelFile.Load(source);
            first.SaveAsExcel(exported);
            var second = ExcelFile.Load(exported);

            await Assert.That(second.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(21.0);
            await Assert.That(second.GetCellValue("Data", "A2").ToDouble()).IsEqualTo(42.0);
            await Assert.That(second.GetCellValue("Data", "A3").ToText()).IsEqualTo("  padded  ");
        }
        finally
        {
            File.Delete(source);
            File.Delete(exported);
        }
    }
}
