using System.IO.Compression;
using System.Text;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Phase 2 audit (shared-formula-delta-production.md, G3 spike follow-up): exhaustive per-function-family
/// coverage for a <see cref="Expressions.SharedFormulaSlave"/> whose master tree passes an
/// <see cref="Expressions.AnchoredCellReference"/>/<see cref="Expressions.AnchoredRangeReference"/> directly
/// to a function that pattern-matches the CONCRETE <c>CellReference</c>/<c>RangeReference</c> shapes — the
/// exact gap the spike's own report flagged (only <c>NumericAggregation</c>/<c>ReferenceGuard</c> were
/// adapted). Every fixture below is deliberately built so the MASTER and SLAVE rows read DIFFERENT underlying
/// data, so a "delta silently not applied" regression shows up as a wrong (usually equal-to-master) value
/// instead of passing by accident.
/// </summary>
public class SharedFormulaSlaveFunctionTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>Single-sheet ("S") raw fixture — mirrors SharedFormulaDeltaTests' builder, trimmed to one sheet
    /// since none of this family's fixtures need a cross-sheet shift.</summary>
    private static async Task WithFixture(string sheetDataXml, Func<Workbook, Task> assert)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-sfsf-{Guid.NewGuid():N}.xlsx");

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
                    {sheetDataXml}
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

    // ROW(ref)/COLUMN(ref): the argument is an AnchoredCellReference inside the master tree. Before the fix,
    // Row/Column's `Arguments switch` only matched the concrete CellReference, so this fell to `_ => #VALUE!`.
    [Test]
    public async Task Row_WithCellArgument_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="B1"><f t="shared" ref="B1:B2" si="0">ROW(A1)</f><v>0</v></c></row>
            <row r="2"><c r="B2"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(1.0);
                // B2 (dr=1): ROW(A1) shifts to ROW(A2) = 2 — not ROW(A1) = 1 again.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(2.0);
            }
        );
    }

    [Test]
    public async Task Column_WithCellArgument_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="B1"><f t="shared" ref="B1:C1" si="0">COLUMN(A1)</f><v>0</v></c><c r="C1"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B1").ToDouble()).IsEqualTo(1.0);
                // C1 (dc=1): COLUMN(A1) shifts to COLUMN(B1) = 2.
                await Assert.That(workbook.GetCellValue("S", "C1").ToDouble()).IsEqualTo(2.0);
            }
        );
    }

    // ROW()/COLUMN() with NO argument: already correct by construction (EvaluationContext.WithDelta leaves
    // SheetName/CellId untouched — see EvaluationContext.cs), pinned here as an explicit regression test
    // together with ADDRESS, which has no reference pattern-matching of its own.
    [Test]
    public async Task Address_WithNoArgumentRowColumn_UsesSlaveOwnCell()
    {
        await WithFixture(
            """
            <row r="2"><c r="B2"><f t="shared" ref="B2:B3" si="0">ADDRESS(ROW(),COLUMN())</f><v>0</v></c></row>
            <row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToText()).IsEqualTo("$B$2");
                await Assert.That(workbook.GetCellValue("S", "B3").ToText()).IsEqualTo("$B$3");
            }
        );
    }

    // OFFSET's base is resolved via NamedReferences.TryResolveReference (the polymorphic path), so this was
    // already correct — pinned as a regression test.
    [Test]
    public async Task Offset_RelativeBase_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">OFFSET(A1,1,0)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): OFFSET(A1,1,0) = A2 = 2.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(2.0);
                // B3 (dr=1): base shifts to A2, so OFFSET(A2,1,0) = A3 = 3.
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(3.0);
            }
        );
    }

    // INDIRECT parses ref_text fresh (never touches the master's own anchored nodes) and ROW() with no
    // argument is already correct — pinned as a regression test for the combined idiom.
    [Test]
    public async Task Indirect_WithRowFunction_UsesSlaveOwnRow()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>11</v></c></row>
            <row r="2"><c r="A2"><v>22</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">INDIRECT("A"&amp;ROW())</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>33</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(22.0);
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(33.0);
            }
        );
    }

    // INDEX's range argument resolves via NamedReferences.TryResolveReference too — already correct — pinned
    // as a regression test with a genuinely relative (non-anchored-with-$) range.
    [Test]
    public async Task Index_RelativeRange_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>10</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>20</v></c><c r="C2"><f t="shared" ref="C2:C3" si="0">INDEX(A1:B3,2,1)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>30</v></c><c r="C3"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="4"><c r="A4"><v>4</v></c><c r="B4"><v>40</v></c></row>
            """,
            async workbook =>
            {
                // C2 (dr=0): INDEX(A1:B3,2,1) = row 2 of A1:B3 = A2 = 2.
                await Assert.That(workbook.GetCellValue("S", "C2").ToDouble()).IsEqualTo(2.0);
                // C3 (dr=1): range shifts to A2:B4, row 2 of THAT = A3 = 3.
                await Assert.That(workbook.GetCellValue("S", "C3").ToDouble()).IsEqualTo(3.0);
            }
        );
    }

    // SUBTOTAL: the fixed bug. Before the fix, an AnchoredRangeReference argument fell to `default` in
    // GatherSkippingSubtotals, which evaluates it directly — AnchoredRangeReference.Evaluate always returns
    // #VALUE! (a range has no scalar value, like RangeReference) — so BOTH cells were #VALUE!.
    [Test]
    public async Task Subtotal_RelativeRange_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">SUBTOTAL(9,A1:A3)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="4"><c r="A4"><v>4</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): SUBTOTAL(9, A1:A3) = SUM(1,2,3) = 6.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(6.0);
                // B3 (dr=1): range shifts to A2:A4 = SUM(2,3,4) = 9.
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(9.0);
            }
        );
    }

    // NPV: the fixed bug. A shared-formula GROUP'S MASTER CELL itself is parsed as an ORDINARY tree with a
    // plain CellReference (see WorksheetStreamLoader.ReadCell: only cells AFTER the master become a
    // SharedFormulaSlave wrapping the anchored tree) — so the divergent value must sit in the SLAVE's own
    // row, not the master's, or the bug never triggers. Here the master (C2) reads a NUMBER (B2=100,
    // always-correct — plain CellReference), the slave (C3, dr=1) reads a TEXT cell (B3="n/a") through its
    // AnchoredCellReference argument: before the fix, that fell to NPV's `default` branch, which uses the
    // DIRECT-value rule (numeric-parses text, fails, #VALUE!) instead of the REFERENCED rule (ignores text).
    [Test]
    public async Task Npv_WithCellArgument_TreatsReferencedTextAsIgnored()
    {
        await WithFixture(
            """
            <row r="2"><c r="B2"><v>100</v></c><c r="C2"><f t="shared" ref="C2:C3" si="0">NPV(0.1,B2)</f><v>0</v></c></row>
            <row r="3"><c r="B3" t="str"><v>n/a</v></c><c r="C3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                // C2 (dr=0, master, plain CellReference): NPV(0.1, B2=100) = 100/1.1.
                await Assert
                    .That(workbook.GetCellValue("S", "C2").ToDouble())
                    .IsEqualTo(100.0 / 1.1)
                    .Within(1e-9);
                // C3 (dr=1, slave, AnchoredCellReference): argument shifts to B3="n/a" — referenced text is
                // IGNORED (not #VALUE!) → sum of nothing = 0.
                await Assert.That(workbook.GetCellValue("S", "C3").ToDouble()).IsEqualTo(0.0);
            }
        );
    }

    // VLOOKUP's table resolves via NamedReferences.TryResolveReference — already correct — pinned as a
    // regression test where BOTH the table and the lookup key are relative, so master/slave search genuinely
    // different (shifted) sub-tables for genuinely different keys.
    [Test]
    public async Task VLookup_RelativeTableAndKey_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>10</v></c><c r="B1" t="str"><v>ten</v></c><c r="C1"><v>20</v></c><c r="D1"><f t="shared" ref="D1:D2" si="0">VLOOKUP(C1,A1:B3,2,FALSE)</f><v>0</v></c></row>
            <row r="2"><c r="A2"><v>20</v></c><c r="B2" t="str"><v>twenty</v></c><c r="C2"><v>40</v></c><c r="D2"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="3"><c r="A3"><v>30</v></c><c r="B3" t="str"><v>thirty</v></c></row>
            <row r="4"><c r="A4"><v>40</v></c><c r="B4" t="str"><v>forty</v></c></row>
            """,
            async workbook =>
            {
                // D1 (dr=0): VLOOKUP(C1=20, A1:B3, 2, FALSE) → table {10 ten; 20 twenty; 30 thirty} → "twenty".
                await Assert.That(workbook.GetCellValue("S", "D1").ToText()).IsEqualTo("twenty");
                // D2 (dr=1): key shifts to C2=40, table shifts to A2:B4 {20 twenty;30 thirty;40 forty} → "forty".
                await Assert.That(workbook.GetCellValue("S", "D2").ToText()).IsEqualTo("forty");
            }
        );
    }

    // SUMIF: the fixed bug (ArgumentFlattening/CriteriaScan.PositionalRange). Before the fix, an
    // AnchoredRangeReference argument fell to `argument.Evaluate(context)` in ArgumentFlattening's default
    // branch, which always yields #VALUE! for a range.
    [Test]
    public async Task SumIf_RelativeRange_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>100</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">SUMIF(A1:A3,"&gt;0")</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="4"><c r="A4"><v>-5</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): SUMIF(A1:A3,">0") over {100,2,3} = 105.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(105.0);
                // B3 (dr=1): range shifts to A2:A4 over {2,3,-5} → only 2+3 = 5 (the negative is excluded, and
                // 100 has scrolled out of the window).
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(5.0);
            }
        );
    }

    // COUNTIF: the fixed bug (RangeValueCursor). Before the fix, an AnchoredRangeReference argument fell to
    // RangeValueCursor.Open's `default` branch, which evaluates it directly and wraps the resulting #VALUE!
    // in a single-element cursor.
    [Test]
    public async Task CountIf_RelativeRange_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>5</v></c></row>
            <row r="2"><c r="A2"><v>5</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">COUNTIF(A1:A3,5)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>1</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="4"><c r="A4"><v>1</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): COUNTIF(A1:A3,5) over {5,5,1} = 2.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(2.0);
                // B3 (dr=1): range shifts to A2:A4 over {5,1,1} = 1.
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(1.0);
            }
        );
    }

    // MATCH shares RangeValueCursor with COUNTIF — same fix, pinned separately since MATCH's exact-match path
    // (matchType=0) is a distinct call site from COUNTIF's.
    [Test]
    public async Task Match_RelativeRange_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>7</v></c></row>
            <row r="2"><c r="A2"><v>8</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">MATCH(9,A1:A3,0)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>9</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="4"><c r="A4"><v>10</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): MATCH(9,A1:A3,0) over {7,8,9} → position 3.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(3.0);
                // B3 (dr=1): range shifts to A2:A4 over {8,9,10} → position 2.
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(2.0);
            }
        );
    }

    // ISFORMULA: the fixed bug — mirrors FormulaText's identical fix in InformationFunctions.cs.
    [Test]
    public async Task IsFormula_WithCellArgument_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><f>1+1</f><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">ISFORMULA(A1)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): ISFORMULA(A1) — A1 is a plain literal → FALSE.
                await Assert.That(workbook.GetCellValue("S", "B2").ToBoolean()).IsFalse();
                // B3 (dr=1): argument shifts to A2, which IS a formula → TRUE.
                await Assert.That(workbook.GetCellValue("S", "B3").ToBoolean()).IsTrue();
            }
        );
    }

    // SHEET(ref): the fixed bug. Before the fix, an AnchoredCellReference argument fell to SheetNumber's
    // `[var argument] => argument.Evaluate(context).AsString()` arm — AnchoredCellReference.Evaluate correctly
    // dereferences to the CELL'S VALUE (not its identity), and ComputedValue.AsString() is null for a numeric
    // value, so SHEET(ref) degraded to #REF! for the slave. A single-sheet fixture cannot make master and
    // slave DIFFER (SheetName is a literal, unaffected by the row/column delta), so this pins correctness
    // (not #REF!) rather than a differing value.
    [Test]
    public async Task Sheet_WithCellArgument_ResolvesInsteadOfErroring()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">SHEET(A1)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(1.0);
                // B3 (dr=1, slave): SHEET(A2) — still sheet "S", position 1, not #REF!.
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(1.0);
            }
        );
    }

    // FORMULATEXT: the fixed bug, direction 1 — a FORMULATEXT(ref) argument INSIDE the shared-formula master.
    [Test]
    public async Task FormulaText_OfAnchoredArgument_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><f>10*2</f><v>20</v></c></row>
            <row r="2"><c r="A2"><f>30*2</f><v>60</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">FORMULATEXT(A1)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><f>40*2</f><v>80</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "B2").ToText()).IsEqualTo("=10*2");
                // B3 (dr=1): argument shifts to A2.
                await Assert.That(workbook.GetCellValue("S", "B3").ToText()).IsEqualTo("=30*2");
            }
        );
    }

    // FORMULATEXT: direction 2 — called from OUTSIDE the shared-formula group, on a cell that IS a
    // SharedFormulaSlave itself. Already correct (FormulaWriter already special-cases the anchored/slave
    // nodes, per the G3 spike) — pinned as a regression test, since this is the "un-parse the slave" half the
    // task explicitly calls out.
    [Test]
    public async Task FormulaText_OfSlaveCell_RendersDeltaShiftedFormula()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">A1*2</f><v>0</v></c><c r="C2"><f>FORMULATEXT(B2)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c><c r="C3"><f>FORMULATEXT(B3)</f><v>0</v></c></row>
            """,
            async workbook =>
            {
                await Assert.That(workbook.GetCellValue("S", "C2").ToText()).IsEqualTo("=A1*2");
                // B3 is the slave (dr=1): its un-parsed formula shows the SHIFTED reference, A2, not A1.
                await Assert.That(workbook.GetCellValue("S", "C3").ToText()).IsEqualTo("=A2*2");
            }
        );
    }

    // ArrayEvaluation's mini-CSE idiom: SMALL(IF(range=…, ROW(range))) needs BOTH the IF-comparison range and
    // the ROW(range) argument to build as arrays over an anchored range. Before the fix, the anchored range
    // fell to ArrayEvaluation's opaque-scalar `default` branch (evaluated once via
    // AnchoredRangeReference.Evaluate, which always yields #VALUE!), so the whole idiom broke.
    [Test]
    public async Task ArrayIdiom_SmallIfRow_OverAnchoredRange_AppliesSlaveDelta()
    {
        await WithFixture(
            """
            <row r="1"><c r="A1"><v>0</v></c></row>
            <row r="2"><c r="A2"><v>1</v></c><c r="B2"><f t="shared" ref="B2:B3" si="0">SMALL(IF(A1:A3=0,ROW(A1:A3)),1)</f><v>0</v></c></row>
            <row r="3"><c r="A3"><v>1</v></c><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>
            <row r="4"><c r="A4"><v>0</v></c></row>
            """,
            async workbook =>
            {
                // B2 (dr=0): over A1:A3 = {0,1,1}, only row 1 (A1) is 0 → SMALL(…,1) = 1.
                await Assert.That(workbook.GetCellValue("S", "B2").ToDouble()).IsEqualTo(1.0);
                // B3 (dr=1): range shifts to A2:A4 = {1,1,0} → only row 4 (A4) is 0 → SMALL(…,1) = 4.
                await Assert.That(workbook.GetCellValue("S", "B3").ToDouble()).IsEqualTo(4.0);
            }
        );
    }
}
