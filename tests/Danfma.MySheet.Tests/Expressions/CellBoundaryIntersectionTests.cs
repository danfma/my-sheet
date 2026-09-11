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

    // === Phase 8: the array half of the boundary is still ABSENT (Phase 7 owns it) ==========================

    [Test]
    public async Task BareLiftedFunction_IsStillValueError_TheArrayHalfBelongsToPhase7()
    {
        // Phase 8 lifts pure-scalar built-ins and unary '-'/'%' over an array — but ONLY inside the mini-CSE,
        // which is entered exclusively by consumers that call ArrayEvaluation themselves (SUM/COUNT/
        // SUMPRODUCT, SMALL/LARGE, INDEX, AGGREGATE). Workbook.EvaluateCell is NOT one of them: it inspects
        // only value.TryGetReference(...) → ImplicitIntersection.Apply, and a lifted function is not a
        // reference — it is Len.Evaluate's untouched #VALUE!.
        //
        // So a bare =LEN(A1:A3) typed in a cell still answers #VALUE! after Phase 8, exactly as =A1:A3*2 and
        // =IF(B2:B5="Show",1,0) do (this class's own comment: the ARRAY half "has no producer in this engine
        // yet and is deliberately absent"). That is a documented GAP, not a rule: the array half means giving
        // EvaluateCell a second arm that runs IsArrayEligible over the WHOLE cell expression and takes
        // ElementAt(0) — a new call site on every cell's hot path — and it belongs to PHASE 7, the phase that
        // creates producers users will type bare. Pinning it here means Phase 7 must change this assertion
        // deliberately instead of discovering the behaviour by accident.
        //
        // THE DIVERGENCE IS IN BOTH ENTRY MODES, AND AN EARLIER VERSION OF THIS COMMENT SAID OTHERWISE. It
        // claimed "=LEN(A1:A3) entered plainly is #VALUE!", which is only true when the formula sits OUTSIDE
        // the range's rows. Re-measured on Aspose.Cells 26.6.0 (2026-09-10), the formula in C1 / C2 / C3 —
        // rows the range covers — and in C5, which it does not, over A1:A3 = 5, 0, 9:
        //
        //     formula            PLAIN C1/C2/C3/C5          ARRAY-ENTERED, every cell
        //     =LEN(A1:A3)        1, 1, 1, #VALUE!           1
        //     =-A1:A3            -5, 0, -9, #VALUE!         -5
        //     =ROUND(A1:A3,0)    5, 0, 9, #VALUE!           5
        //     =A1:A3*2           10, 0, 18, #VALUE!         10
        //
        // Plain entry intersects PER ROW (which is why =-A1:A3 walks -5, 0, -9 while LEN coincidentally
        // answers 1 on all three single-digit cells), and array entry answers the top-left. MySheet answers
        // #VALUE! in every one of those cells, so the gap is a divergence from BOTH oracle columns, not
        // parity with the plain one. The assertions below use C2, an INSIDE row, so they pin the divergence
        // at its widest: the oracle says 1 there.
        await Assert.That(InCell("C2", "=LEN(A1:A3)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InCell("C2", "=-A1:A3")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InCell("C2", "=ROUND(A1:A3,0)")).IsEqualTo(ErrorValue.NotValue);

        // The RANGE half — Phase 1's FIX B, commit b89b856 — DID land and is unaffected: the same cell, same
        // row, holding the bare range intersects to A2 = 0 (and row 3 to A3 = 9, so the 0 is not a blank
        // coercion in disguise). The contrast is the whole point: the boundary knows references, not arrays.
        await Assert.That(InCell("C2", "=A1:A3")).IsEqualTo(0.0);
        await Assert.That(InCell("C3", "=A1:A3")).IsEqualTo(9.0);

        // And a CONSUMED lift in the same cell is fine — the consumer enters the mini-CSE itself.
        // LEN(5)+LEN(0)+LEN(9) = 3 (Aspose 26.6.0, CSE-entered, 2026-09-09).
        await Assert.That(InCell("C2", "=SUM(LEN(A1:A3))")).IsEqualTo(3.0);
    }

    // === Phase 10: a BROADCAST product at the boundary is still absent, for the same reason ==============

    [Test]
    public async Task BareBroadcastProduct_IsStillValueError_TheArrayHalfBelongsToPhase7()
    {
        // Phase 10 taught the mini-CSE to broadcast mismatched vectors — but, like Phase 8's lifting, only
        // INSIDE the mini-CSE, which Workbook.EvaluateCell does not enter. So a bare product typed in a
        // cell is still the multiplication operator's own #VALUE!, unchanged by this phase.
        //
        // This is an UNCHANGED-BY-DESIGN pin over a DIVERGENCE, and the divergence is what makes it worth
        // writing down. Aspose.Cells 26.6.0, measured 2026-09-10, PLAIN entry — the only column that
        // exists here, since a formula typed into a cell has no CSE twin:
        //
        //   J1: =A1:A3*E1:E3  →  1        J3: =A1:A3*E1:E3  →  21       (per-OPERAND implicit intersection,
        //   J2: =A1:C3*E1:E3  →  #VALUE!  L2: =E1:E3*10     →  20        each operand taken on the formula's
        //                                                                own row: A3*E3 = 7*3, E3*10 = 3*10)
        //
        // The J1/J3 pair is what identifies the oracle's mechanism: 1 and 21 from the SAME formula on two
        // rows is legacy per-operand intersection, not S4's "top-left element of a computed array" (which
        // would give 1 on both rows). Phase 7 owns the array half of the boundary and must reconcile S4
        // with this measurement; pinning today's #VALUE! here means that reconciliation edits these lines
        // deliberately instead of moving the answer by accident.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        var value = 1;
        foreach (var row in new[] { 1, 2, 3 })
        {
            foreach (var column in new[] { "A", "B", "C" })
            {
                sheet[$"{column}{row}"] = Number(value++);
            }
        }

        sheet["E1"] = Number(1);
        sheet["E2"] = Number(2);
        sheet["E3"] = Number(3);

        object? InGridCell(string id, string formula)
        {
            sheet[id] = ExpressionParser.Parse(formula, sheet);

            return workbook.GetCellValue("Sheet1", id).AsObject();
        }

        await Assert.That(InGridCell("J3", "=A1:A3*E1:E3")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InGridCell("J1", "=A1:A3*E1:E3")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InGridCell("J2", "=A1:C3*E1:E3")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(InGridCell("L2", "=E1:E3*10")).IsEqualTo(ErrorValue.NotValue);

        // The counterweight, and the reason the gap is in the BOUNDARY and not in the phase: the very same
        // product CONSUMED in the very same cell broadcasts and answers — 6*1 + 15*2 + 24*3 (Aspose.Cells
        // 26.6.0, 2026-09-10, CSE column; VectorBroadcastingTests owns the family of that number).
        await Assert.That(InGridCell("L4", "=SUM(A1:C3*E1:E3)")).IsEqualTo(108.0);
    }
}
