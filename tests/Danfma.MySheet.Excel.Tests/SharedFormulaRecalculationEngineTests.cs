using System.IO.Compression;
using System.Text;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Fase 3 (shared-formula-delta): the engine gap this closes — before this test existed, no test in the repo
/// exercised <see cref="RecalculationEngine"/> against a workbook loaded from a real .xlsx with shared
/// formulas (i.e. real <c>SharedFormulaSlave</c> cells, not ones built by hand in a unit test). It doubles as
/// the strongest possible regression guard for <c>DependencyExtractor</c>'s delta-aware rewrite (see
/// <c>SharedFormulaDependencyParityTests</c> in the core test project): before that fix, every
/// <c>SharedFormulaSlave</c> cell was <c>AlwaysDirty</c>, and <see cref="DirtyGraph.DirtyEngine.CalculateDirty"/>
/// unions the always-dirty set into EVERY <see cref="RecalculationEngine.Recalculate"/> call unconditionally —
/// so editing ONE unrelated input would have evicted every shared-formula slave in the workbook, not just the
/// ones that actually depend on it. <see cref="RecalculationResult.DirtyCellCount"/> is the precise, publicly
/// observable signal: it must stay tiny (edited cell + its one true dependent) even though the fixture's
/// shared-formula groups hold several slaves.
/// </summary>
public class SharedFormulaRecalculationEngineTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    // Single-sheet fixture with TWO shared-formula groups sharing the same grid:
    //  - scalar group (si=0): master B2="A2*2" ref B2:B6 — slaves B3..B6 (deltaRow 1..4), each reading a
    //    DIFFERENT single input cell (A3..A6).
    //  - range group (si=1): master G2="SUM(D2:F2)" ref G2:G4 — slaves G3,G4 (deltaRow 1,2), each reading a
    //    DIFFERENT 3-cell row (D..F of its own row).
    // Values: A2..A6 = 2..6 so B2..B6 = 4,6,8,10,12; D/E/F rows 2..4 = 10/20/30, 100/200/300, 1000/2000/3000
    // so G2=60, G3=600, G4=6000.
    private static Workbook LoadFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-sfre-{Guid.NewGuid():N}.xlsx");

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
                    <row r="2">
                    <c r="A2"><v>2</v></c>
                    <c r="B2"><f t="shared" ref="B2:B6" si="0">A2*2</f><v>0</v></c>
                    <c r="D2"><v>10</v></c><c r="E2"><v>20</v></c><c r="F2"><v>30</v></c>
                    <c r="G2"><f t="shared" ref="G2:G4" si="1">SUM(D2:F2)</f><v>0</v></c>
                    </row>
                    <row r="3">
                    <c r="A3"><v>3</v></c>
                    <c r="B3"><f t="shared" si="0"/><v>0</v></c>
                    <c r="D3"><v>100</v></c><c r="E3"><v>200</v></c><c r="F3"><v>300</v></c>
                    <c r="G3"><f t="shared" si="1"/><v>0</v></c>
                    </row>
                    <row r="4">
                    <c r="A4"><v>4</v></c>
                    <c r="B4"><f t="shared" si="0"/><v>0</v></c>
                    <c r="D4"><v>1000</v></c><c r="E4"><v>2000</v></c><c r="F4"><v>3000</v></c>
                    <c r="G4"><f t="shared" si="1"/><v>0</v></c>
                    </row>
                    <row r="5">
                    <c r="A5"><v>5</v></c>
                    <c r="B5"><f t="shared" si="0"/><v>0</v></c>
                    </row>
                    <row r="6">
                    <c r="A6"><v>6</v></c>
                    <c r="B6"><f t="shared" si="0"/><v>0</v></c>
                    </row>
                    </sheetData>
                    </worksheet>
                    """
                );
            }

            return ExcelFile.Load(path);
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

    private static CellRef Ref(string id) => new("S", id);

    // Warms the memoized cache for every formula cell in the fixture, mirroring the engine's documented
    // contract ("create AFTER the workbook is populated, typically after a first ComputeAll").
    private static async Task<Workbook> WarmedFixture()
    {
        var workbook = LoadFixture();

        await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(4.0);
        await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(6.0);
        await Assert.That(workbook.GetCellValue("S", "B4").ToDouble()).IsEqualTo(8.0);
        await Assert.That(workbook.GetCellValue("S", "B5").ToDouble()).IsEqualTo(10.0);
        await Assert.That(workbook.GetCellValue("S", "B6").ToDouble()).IsEqualTo(12.0);
        await Assert.That(workbook.GetCellValue("S", "G2").ToDouble()).IsEqualTo(60.0);
        await Assert.That(workbook.GetCellValue("S", "G3").ToDouble()).IsEqualTo(600.0);
        await Assert.That(workbook.GetCellValue("S", "G4").ToDouble()).IsEqualTo(6000.0);

        return workbook;
    }

    [Test]
    public async Task ScalarInputEdit_RecomputesOnlyItsOwnSlave_PreciseCone()
    {
        var workbook = await WarmedFixture();
        var engine = workbook.CreateRecalculationEngine();

        workbook.Sheets["S"]["A4"] = new Expressions.NumberValue(40);
        var result = engine.Recalculate([Ref("A4")]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.Partial);
        await Assert.That(result.StructureRebuilt).IsFalse();
        // Precise cone: only A4 (edited) + B4 (its one true dependent) — NOT every slave in the B group.
        await Assert.That(result.DirtyCellCount).IsEqualTo(2);

        await Assert.That(workbook.GetCellValue("S", "B4").ToDouble()).IsEqualTo(80.0);
        // Every other slave/master in both groups is untouched by this edit.
        await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(4.0);
        await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(6.0);
        await Assert.That(workbook.GetCellValue("S", "B5").ToDouble()).IsEqualTo(10.0);
        await Assert.That(workbook.GetCellValue("S", "B6").ToDouble()).IsEqualTo(12.0);
        await Assert.That(workbook.GetCellValue("S", "G2").ToDouble()).IsEqualTo(60.0);
        await Assert.That(workbook.GetCellValue("S", "G3").ToDouble()).IsEqualTo(600.0);
        await Assert.That(workbook.GetCellValue("S", "G4").ToDouble()).IsEqualTo(6000.0);
    }

    [Test]
    public async Task RangeInputEdit_RecomputesOnlyItsOwnSlave_PreciseCone()
    {
        var workbook = await WarmedFixture();
        var engine = workbook.CreateRecalculationEngine();

        workbook.Sheets["S"]["E3"] = new Expressions.NumberValue(2000);
        var result = engine.Recalculate([Ref("E3")]);

        await Assert.That(result.Mode).IsEqualTo(RecalculationMode.Partial);
        await Assert.That(result.StructureRebuilt).IsFalse();
        // Precise cone: only E3 (edited) + G3 (the one slave whose anchored range covers row 3) — NOT G2/G4.
        await Assert.That(result.DirtyCellCount).IsEqualTo(2);

        await Assert.That(workbook.GetCellValue("S", "G3").ToDouble()).IsEqualTo(2400.0); // 100+2000+300
        await Assert.That(workbook.GetCellValue("S", "G2").ToDouble()).IsEqualTo(60.0);
        await Assert.That(workbook.GetCellValue("S", "G4").ToDouble()).IsEqualTo(6000.0);
        await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(4.0);
        await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(6.0);
        await Assert.That(workbook.GetCellValue("S", "B4").ToDouble()).IsEqualTo(8.0);
    }

    [Test]
    public async Task MasterFormulaEdit_RebuildsGraph_AndRecomputesTheMasterCorrectly()
    {
        var workbook = await WarmedFixture();
        var engine = workbook.CreateRecalculationEngine();

        var sheet = workbook.Sheets["S"];
        sheet["B2"] = ExpressionParser.Parse("=A2*3", sheet); // structural: B2's formula shape changed

        var result = engine.Recalculate([Ref("B2")]);

        await Assert.That(result.StructureRebuilt).IsTrue();
        await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(6.0); // 2*3

        // The slaves each captured the group's ORIGINAL anchored tree at load time — a later, independent
        // edit to the master CELL (through the ordinary Sheet API, not the .xlsx loader) does not retroactively
        // change what they evaluate; they keep reading their own frozen shared tree, unaffected.
        await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(6.0);
        await Assert.That(workbook.GetCellValue("S", "B4").ToDouble()).IsEqualTo(8.0);
        await Assert.That(workbook.GetCellValue("S", "B5").ToDouble()).IsEqualTo(10.0);
        await Assert.That(workbook.GetCellValue("S", "B6").ToDouble()).IsEqualTo(12.0);
    }
}
