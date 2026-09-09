using System.Buffers.Binary;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Excel's implicit intersection at the CELL boundary: when a formula's FINAL value is a multi-cell range,
/// the cell shows the range's cell on the formula cell's own row/column (Microsoft's "@" operator rule)
/// instead of <c>#VALUE!</c>.
///
/// <para>The distinguishing feature of every case here is the READ PATH: the formula is stored in a real
/// cell and read back through <see cref="Workbook.GetCellValue(string,string)"/>, the only path that crosses
/// <c>Workbook.EvaluateCell</c>. A <c>Parse(f, sheet).Evaluate(workbook)</c> — what
/// <c>IndirectTests.Calc</c> and most sibling suites do — never crosses that boundary and therefore still
/// sees the bare range's <c>#VALUE!</c> (pinned by
/// <see cref="DirectEvaluate_OfABareRange_IsStillValueError"/>). That is exactly why no pre-existing test
/// covered this, and exactly the trap the next reader will fall into.</para>
/// </summary>
public class CellBoundaryIntersectionTests
{
    // Sheet1 fixture. A2 is a deliberate 0 (a real value that the blank→0 boundary coercion also produces),
    // so every case whose answer is 0 is paired with a discriminating row-3 case whose answer is 9. Sheet2's
    // column A holds 11/22/33 — deliberate lies, so a cross-sheet intersection that read the WRONG sheet
    // would be visible.
    private static Workbook Fixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = Number(5);
        sheet["A2"] = Number(0);
        sheet["A3"] = Number(9);
        sheet["B1"] = Number(2);
        sheet["D1"] = String("txt");
        sheet["E1"] = ExpressionParser.Parse("=1/0", sheet);

        var second = workbook.Sheets.Add("Sheet2");
        second["A1"] = Number(11);
        second["A2"] = Number(22);
        second["A3"] = Number(33);

        workbook.DefineName("MyName", "Sheet1!$A$1:$A$3");
        workbook.DefineName("Ghosty", "Ghost!$A$1:$A$3");

        return workbook;
    }

    // Stores `formula` at `id` on a pristine fixture and reads it back through the cell, so the value is the
    // one a HOST sees. A fresh workbook per call keeps the probe cells from colliding across cases.
    private static object? InCell(string id, string formula, string sheetName = "Sheet1")
    {
        var workbook = Fixture();
        var sheet = workbook[sheetName];
        sheet[id] = ExpressionParser.Parse(formula, sheet);

        return workbook.GetCellValue(sheetName, id).AsObject();
    }

    // === Closed ranges ===================================================================================

    [Test]
    public async Task BareSingleColumnRange_IntersectsTheFormulaRow()
    {
        await Assert.That(InCell("C2", "=A1:A3")).IsEqualTo(0.0);
        await Assert.That(InCell("C3", "=A1:A3")).IsEqualTo(9.0);
    }

    [Test]
    public async Task BareRange_WhenTheFormulaRowIsOutsideIt_IsValueError()
    {
        await Assert.That(InCell("C9", "=A1:A3")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task BareRange_IntersectingTheFormulaCellItself_IsReferenceError()
    {
        // A2 holding =A1:A3 intersects A2 — itself. The deref re-enters the cell that is already on the
        // evaluation stack, so the cycle guard answers #REF! (Excel raises a circular-reference dialog).
        await Assert.That(InCell("A2", "=A1:A3")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task BareSingleRowRange_IntersectsTheFormulaColumn()
    {
        await Assert.That(InCell("B5", "=A1:C1")).IsEqualTo(2.0);
    }

    [Test]
    public async Task RangeBounds_AreInclusiveOnEveryEdge()
    {
        // The containment test is >= / <= on both axes, and only an EDGE case can tell that from > / <. Of
        // the four edges the cases above pin exactly one (C3 of A1:A3 is the BOTTOM row); the rest sit
        // strictly inside the span (C2 of A1:A3, B5's middle column of A1:C1). One line per edge here, each
        // discriminating against the #VALUE! an exclusive bound would give, with the bottom row repeated so
        // the rule reads as one piece.
        //
        // Rows of the single-COLUMN range A1:A3, from the formula cell's own row: TOP then BOTTOM.
        await Assert.That(InCell("C1", "=A1:A3")).IsEqualTo(5.0);
        await Assert.That(InCell("C3", "=A1:A3")).IsEqualTo(9.0);

        // Columns of the single-ROW range A1:C1, from the formula cell's own column: LEFT (column A) and
        // RIGHT (column C). C1 is EMPTY in the fixture, so the right edge shows the boundary's blank→0 — a 0
        // that still separates it from #VALUE! (an exclusive right bound), from A1's 5, from B1's 2 and from
        // D1's "txt" (a bound one column too wide).
        await Assert.That(InCell("A5", "=A1:C1")).IsEqualTo(5.0);
        await Assert.That(InCell("C5", "=A1:C1")).IsEqualTo(0.0);
    }

    [Test]
    public async Task BareTwoDimensionalRange_IsValueError()
    {
        // Both axes span more than one cell: the @ operator has no single answer, so #VALUE! stands.
        await Assert.That(InCell("B2", "=A1:C3")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InCell("C1", "=A1:C3")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task OneByOneRange_IsItsOwnCell_RegardlessOfPosition()
    {
        await Assert.That(InCell("Z99", "=A1:A1")).IsEqualTo(5.0);
    }

    // === Open ranges (DECLARED bounds, not populated) ====================================================

    [Test]
    public async Task WholeColumn_IntersectsTheFormulaRow_EvenWhenThatCellIsEmpty()
    {
        // A7 is empty: the DECLARED row axis of A:A is unbounded, so row 7 is inside and the blank becomes
        // the boundary's 0. The POPULATED box (A1:A3) would have answered #VALUE! here.
        await Assert.That(InCell("C7", "=A:A")).IsEqualTo(0.0);
    }

    [Test]
    public async Task WholeRow_IntersectsTheFormulaColumn()
    {
        await Assert.That(InCell("C7", "=1:1")).IsEqualTo(0.0);
        await Assert.That(InCell("B7", "=1:1")).IsEqualTo(2.0);
    }

    [Test]
    public async Task OneSidedOpenRange_HonoursTheKnownBound()
    {
        await Assert.That(InCell("C1", "=A2:A")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InCell("C3", "=A2:A")).IsEqualTo(9.0);
        await Assert.That(InCell("C77", "=A:A10")).IsEqualTo(ErrorValue.NotValue);
    }

    // === Names, sheets, unions ============================================================================

    [Test]
    public async Task DefinedNameOverARange_Intersects()
    {
        await Assert.That(InCell("C2", "=MyName")).IsEqualTo(0.0);
        await Assert.That(InCell("C3", "=MyName")).IsEqualTo(9.0);
    }

    [Test]
    public async Task CrossSheetRange_IntersectsPositionally()
    {
        // Typed in Sheet2!C2, =Sheet1!A1:A3 yields Sheet1!A2 (0) — the formula cell's ROW/COLUMN, not its
        // sheet, drives the intersection. Sheet2!A2 is 22, so reading the wrong sheet would show.
        await Assert.That(InCell("C2", "=Sheet1!A1:A3", "Sheet2")).IsEqualTo(0.0);
    }

    [Test]
    public async Task RangeOnAMissingSheet_IsReferenceError()
    {
        await Assert.That(InCell("C2", "=Ghost!A1:A3")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(InCell("C2", "=Ghosty")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Union_StaysValueError()
    {
        // A union of areas has no row/column axis to intersect: deliberately unchanged.
        await Assert.That(InCell("C2", "=(A1:A3,B1:B3)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task HostBuiltRange_WithANonA1CornerId_IsValueError()
    {
        // A RangeReference the PARSER can never build: a host assembled the node itself and gave a corner an
        // id that is not an A1 address. Before the boundary intersection such a cell answered #VALUE! from
        // RangeReference.Evaluate WITHOUT parsing either corner, so the intersection must not turn it into a
        // FormatException escaping Workbook.GetCellValue — a value-returning public API.
        var workbook = Fixture();
        workbook["Sheet1"]["C2"] = new RangeReference("bogus", "A3", "Sheet1");

        await Assert
            .That(workbook.GetCellValue("Sheet1", "C2").AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Intersection_YieldsTheTargetCellsOwnKindAndErrors()
    {
        await Assert.That(InCell("C1", "=D1:D3")).IsEqualTo("txt");
        await Assert.That(InCell("C1", "=E1:E3")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task ComputedArrayInACell_StaysValueError_WhereABareRangeIntersects()
    {
        // The boundary intersects a REFERENCE, and a computed array is not one: IF returns the taken
        // branch's own Evaluate — RangeReference.Evaluate's #VALUE! — so NamedReferences.CaptureValue (which
        // only looks at the TOP node) never sees a reference and there is nothing to intersect. The array
        // half of Excel's rule, which would collapse a computed array to its top-left value, has no producer
        // here and is deliberately absent (see ImplicitIntersection's summary).
        //
        // MiniCseConsumerTests.DryCell_IfArray_StaysValueError pins the same rule on the direct
        // Expression.Evaluate path, which never crosses the cell boundary; this is the CELL path, the only
        // one where an intersection could have happened.
        await Assert.That(InCell("C3", "=IF(TRUE,A1:A3,B1)")).IsEqualTo(ErrorValue.NotValue);

        // The same formula CELL, so the two answers differ by the node kind alone and not by position: the
        // bare range there intersects to A3.
        await Assert.That(InCell("C3", "=A1:A3")).IsEqualTo(9.0);
    }

    // === Every producer of a reference-kind final value ==================================================

    [Test]
    [Arguments("=INDIRECT(\"MyName\")")]
    [Arguments("=INDIRECT(\"A1:A3\")")]
    [Arguments("=OFFSET(A1,0,0,3,1)")]
    [Arguments("=CHOOSE(1,A1:A3)")]
    [Arguments("=+A1:A3")]
    [Arguments("=LET(x,A1:A3,x)")]
    [Arguments("=INDEX(A1:A3,2,1):A3")]
    public async Task ReferenceProducingFormula_IntersectsAtTheCell(string formula)
    {
        await Assert.That(InCell("C2", formula)).IsEqualTo(0.0);
        await Assert.That(InCell("C3", formula)).IsEqualTo(9.0);
    }

    [Test]
    public async Task SharedFormulaGroupOverABareRange_IntersectsAtEachSlavesOwnCell()
    {
        // The master's body is a bare anchored range; the slave two rows down evaluates the same shared tree
        // with a +2 row delta, so it denotes A3:A5 and intersects at ITS own row (3) — the master, at row 1,
        // intersects A1. A slave over a defined name resolves to the same rectangle for every slave and so
        // intersects per row as well.
        var workbook = Fixture();
        var sheet = workbook["Sheet1"];
        var master = ExpressionParser.ParseAnchoredMasterBody(
            ExpressionParser.TokenizeFormulaBody("A1:A3"),
            sheet
        );

        sheet["C1"] = master;
        sheet["C3"] = new SharedFormulaSlave(master, 2, 0);
        sheet["D3"] = new SharedFormulaSlave(new NameReference("MyName"), 2, 0);

        await Assert.That(workbook.GetCellValue("Sheet1", "C1").AsObject()).IsEqualTo(5.0);
        await Assert.That(workbook.GetCellValue("Sheet1", "C3").AsObject()).IsEqualTo(9.0);
        await Assert.That(workbook.GetCellValue("Sheet1", "D3").AsObject()).IsEqualTo(9.0);
    }

    // === Unchanged controls ===============================================================================

    [Test]
    public async Task ConsumedInsideAFunction_AndSingleCells_AreUnchanged()
    {
        await Assert.That(InCell("C2", "=SUM(A1:A3)")).IsEqualTo(14.0);
        await Assert.That(InCell("C2", "=A1")).IsEqualTo(5.0);
        await Assert.That(Fixture().GetCellValue("Sheet1", "C5").AsObject()).IsNull();
    }

    [Test]
    public async Task DirectEvaluate_OfABareRange_IsStillValueError()
    {
        // The honest boundary of the change: a range node has no scalar value, and only the CELL applies the
        // implicit intersection. Do NOT "fix" RangeReference.Evaluate instead.
        var workbook = Fixture();

        await Assert
            .That(
                ExpressionParser.Parse("=A1:A3", workbook["Sheet1"]).Evaluate(workbook).AsObject()
            )
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task WarmStart_RoundTripsTheIntersectedValue()
    {
        var workbook = Fixture();
        var sheet = workbook["Sheet1"];
        sheet["C2"] = ExpressionParser.Parse("=MyName", sheet);

        await Assert.That(workbook.GetCellValue("Sheet1", "C2").AsObject()).IsEqualTo(0.0);

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path, new WorkbookSaveOptions { IncludeComputedValues = true });
            var warm = File.ReadAllBytes(path);

            // The intersected value is now IN the snapshot, not merely recomputed after the load: as a
            // reference-kind value C2 used to be dropped by CachedCellValue.TryFrom, and a recompute would
            // also answer 0 — so the value block is what makes this assertion able to fail.
            var modelLength = BinaryPrimitives.ReadInt32LittleEndian(warm.AsSpan(5, 4));
            var values =
                MemoryPackSerializer.Deserialize<List<CachedCellValue>>(
                    warm.AsSpan(9 + modelLength)
                ) ?? [];

            await Assert
                .That(values.Any(entry => entry.SheetName == "Sheet1" && entry.CellId == "C2"))
                .IsTrue();

            var loaded = Workbook.Load(path);

            await Assert.That(loaded.GetCellValue("Sheet1", "C2").AsObject()).IsEqualTo(0.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
