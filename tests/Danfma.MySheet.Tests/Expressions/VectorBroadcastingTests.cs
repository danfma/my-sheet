using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 10 of <c>plans/structured-table-references-and-aggregate.md</c> — VECTOR BROADCASTING in the
/// mini-CSE (<see cref="ArrayEvaluation"/>). Before this phase two array operands had to have EQUAL shapes:
/// a mismatched pair took the per-axis maximum and every operand then refused that extent with a per-ELEMENT
/// <c>#VALUE!</c>, so <c>SUM(A1:C3*E1:E3)</c> was <c>#VALUE!</c> and <c>COUNT(A1:C3*E1:E3)</c> was 0 (both
/// measured on <c>897affc</c>, 2026-09-10) — a per-element fill, not a refusal of the whole expression. Only a
/// scalar broadcast.
///
/// <para>The rule these pins describe, per axis independently: an extent of 1 defers to the other operand (a
/// 1x1 range broadcasts everywhere, a 1xN row repeats down every row, an Nx1 column repeats across every
/// column, and Nx1 against 1xM is the NxM outer product); where BOTH extents exceed 1 and differ, the result
/// takes the LARGER and every position the shorter operand does not cover is <c>#N/A</c> — a per-element
/// error, so <c>COUNT</c> still counts the covered ones, <c>IFERROR</c> recovers them and
/// <c>AGGREGATE(15,6,…)</c> skips them. Error precedence is the operator's existing left-then-right rule: the
/// left operand's own error beats the right operand's uncovered position, and vice versa.</para>
///
/// <para>These are the phase's TDD pins, written BEFORE the engine change: every assertion whose comment says
/// "today" names the value this tree answered when the pin was written (measured on
/// <c>feat/vector-broadcasting</c> at <c>897affc</c>, 2026-09-10), so the fix is a deliberate change of answer
/// rather than a silent one.</para>
///
/// <para>Oracle. Every golden below was MEASURED against the designated P0 oracle, Aspose.Cells 26.6.0
/// (2026-09-10), with each formula entered BOTH plainly and CSE-entered (<c>SetArrayFormula</c>). The CSE
/// column is the oracle for this file: Aspose's PLAIN entry applies legacy implicit intersection inside the
/// argument, so <c>SUM(A1:C3*E1:E3)</c> is <c>#VALUE!</c> there and 108 CSE-entered, and the mini-CSE
/// reproduces CSE semantics by design (<c>plans/mini-cse-array-arguments.md</c>) — it is entered only from a
/// consumed position (SUM/COUNT/SUMPRODUCT/SMALL/INDEX/AGGREGATE), never at the cell boundary. Where the two
/// columns differ and the difference is informative, the plain value is named in the comment as such.</para>
/// </summary>
public class VectorBroadcastingTests
{
    // Grid(): A1:C3 = 1..9 ROW-MAJOR (A1=1, B1=2, C1=3, A2=4, …, C3=9), E1:E3 = 1,2,3 (a 3x1 column),
    // E5:G5 = 10,20,30 (a 1x3 row), H1:H2 = 1,2 (a 2x1 column) and nothing else.
    //
    // The four shapes are what make the rule's clauses distinguishable: E1:E3 pairs with A1:C3's ROWS, E5:G5
    // with its COLUMNS (so 108 and 960 cannot be produced by the same wrong projection), E1:E1/E5:E5 are 1x1
    // RANGES rather than scalars, and H1:H2 is the only operand that is neither 1 nor 3 long — the one that
    // leaves positions uncovered.
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        var value = 1;
        foreach (var row in new[] { 1, 2, 3 })
        {
            foreach (var column in new[] { "A", "B", "C" })
            {
                sheet[$"{column}{row}"] = new NumberValue(value++);
            }
        }

        sheet["E1"] = new NumberValue(1);
        sheet["E2"] = new NumberValue(2);
        sheet["E3"] = new NumberValue(3);
        sheet["E5"] = new NumberValue(10);
        sheet["F5"] = new NumberValue(20);
        sheet["G5"] = new NumberValue(30);
        sheet["H1"] = new NumberValue(1);
        sheet["H2"] = new NumberValue(2);

        return (workbook, sheet);
    }

    private static object? OnGrid(string formula)
    {
        var (workbook, sheet) = Grid();

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Grid() with A3 = =1/0: the LEFT operand's own error sits at (3,1), exactly where H1:H2 leaves the
    // position uncovered, so the two candidate errors compete at one position.
    private static object? OnGridWithErrorInA3(string formula)
    {
        var (workbook, sheet) = Grid();
        sheet["A3"] = ExpressionParser.Parse("=1/0", sheet);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Grid() with B2 = =1/0 and E2 CLEARED: the error is at (2,2) — BEFORE row 3 in row-major scan order —
    // and E1:E3 now carries a BLANK element, which coerces to 0 rather than refusing.
    private static object? OnGridWithErrorInB2(string formula)
    {
        var (workbook, sheet) = Grid();
        sheet["B2"] = ExpressionParser.Parse("=1/0", sheet);
        sheet.Remove("E2");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // --- The broadcasting clauses: 1x1, column, row, outer product ---

    [Test]
    public async Task ColumnVector_RepeatsAcrossEveryColumn()
    {
        // A 3x1 column against a 3x3 rectangle: row r's single value multiplies every column of row r.
        // 6*1 + 15*2 + 24*3 = 108. Today #VALUE!.
        await Assert.That(Num(OnGrid("=SUM(A1:C3*E1:E3)"))).IsEqualTo(108.0);
    }

    [Test]
    public async Task RowVector_RepeatsDownEveryRow()
    {
        // A 1x3 row against the same 3x3: column c's single value multiplies every row of column c.
        // 12*10 + 15*20 + 18*30 = 960 — a different number from the column case above, which is what makes
        // the two clauses independently pinned. Today #VALUE!.
        await Assert.That(Num(OnGrid("=SUM(A1:C3*E5:G5)"))).IsEqualTo(960.0);
    }

    [Test]
    public async Task ColumnAgainstRow_IsTheOuterProduct()
    {
        // 3x1 against 1x3 is a 3x3 result, (1+2+3) * (10+20+30) = 360, and the operand ORDER does not change
        // it. Today #VALUE! both ways.
        await Assert.That(Num(OnGrid("=SUM(E1:E3*E5:G5)"))).IsEqualTo(360.0);
        await Assert.That(Num(OnGrid("=SUM(E5:G5*E1:E3)"))).IsEqualTo(360.0);

        // The same clause with the extents the other way round: 3x1 against 1x2 is a 3x2 result, so COUNT
        // sees SIX elements, not three. Today #VALUE! / 0.
        await Assert.That(Num(OnGrid("=SUM(A1:A3*E5:F5)"))).IsEqualTo(360.0);
        await Assert.That(Num(OnGrid("=COUNT(A1:A3*E5:F5)"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task OneByOneRange_BroadcastsEverywhere_LikeAScalar()
    {
        // A 1x1 RANGE is a vector on both axes, so it broadcasts to all nine positions: 45*1 and 45*10.
        // Today #VALUE! — only a ScalarOperand broadcast, and E1:E1 is a RangeOperand.
        await Assert.That(Num(OnGrid("=SUM(A1:C3*E1:E1)"))).IsEqualTo(45.0);
        await Assert.That(Num(OnGrid("=SUM(A1:C3*E5:E5)"))).IsEqualTo(450.0);
    }

    // --- Both extents above 1 and different: the larger extent, #N/A in the uncovered positions ---

    [Test]
    public async Task MismatchedExtents_TakeTheLargerAndMarkTheUncoveredTailNotAvailable()
    {
        // The four shapes, each SUM → #N/A (the first uncovered position propagates) and each COUNT → the
        // number of COVERED positions. Today #VALUE! / 0 for all eight: the per-element fill left nothing
        // countable.
        //
        // 3x3 against 2x1 → 3x3, row 3 uncovered: 1,2,3 / 8,10,12 / #N/A,#N/A,#N/A.
        await Assert.That(OnGrid("=SUM(A1:C3*H1:H2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(A1:C3*H1:H2)"))).IsEqualTo(6.0);

        // 3x1 against 2x1 → 3x1, row 3 uncovered: 1, 8, #N/A.
        await Assert.That(OnGrid("=SUM(A1:A3*H1:H2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(A1:A3*H1:H2)"))).IsEqualTo(2.0);

        // 3x2 against 2x3 → 3x3: BOTH operands leave positions uncovered — the left has no column 3, the
        // right no row 3 — so only (1,1), (1,2), (2,1), (2,2) = 1, 4, 16, 25 survive.
        await Assert.That(OnGrid("=SUM(A1:B3*A1:C2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(A1:B3*A1:C2)"))).IsEqualTo(4.0);

        // 3x1 against 2x2 → 3x2: the column broadcasts across both columns, the 2x2 leaves row 3 uncovered.
        await Assert.That(OnGrid("=SUM(A1:A3*A1:B2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(A1:A3*A1:B2)"))).IsEqualTo(4.0);
    }

    [Test]
    public async Task TheUncoveredPositions_AreExactlyWhereIndexSaysTheyAre()
    {
        // COUNT can only say HOW MANY positions are covered; INDEX says WHICH, and that is the assertion
        // that separates "the larger extent with #N/A in the uncovered tail" from any other fill of the same
        // cardinality. Today every one of these is #VALUE!.
        //
        // 3x3 against 2x1: the whole of row 3 is uncovered, and nothing else is.
        await Assert.That(OnGrid("=INDEX(A1:C3*H1:H2,3,1)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=INDEX(A1:C3*H1:H2,3,2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=INDEX(A1:C3*H1:H2,3,3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=INDEX(A1:C3*H1:H2,2,1)"))).IsEqualTo(8.0);
        await Assert.That(Num(OnGrid("=INDEX(A1:C3*H1:H2,2,3)"))).IsEqualTo(12.0);

        // 3x2 against 2x3: a COLUMN is uncovered on the left and a ROW on the right, so (1,3) and (3,1) are
        // both #N/A while (2,2) is 5*5 — the case where a row-only or column-only fill would be wrong.
        await Assert.That(OnGrid("=INDEX(A1:B3*A1:C2,1,3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=INDEX(A1:B3*A1:C2,3,1)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=INDEX(A1:B3*A1:C2,2,2)"))).IsEqualTo(25.0);

        // 3x1 against 2x1: the tail is one element long.
        await Assert.That(OnGrid("=INDEX(A1:A3*H1:H2,3,1)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=INDEX(A1:A3*H1:H2,2,1)"))).IsEqualTo(8.0);
    }

    [Test]
    public async Task Index_OverABroadcastArray_ReadsTheBroadcastElement()
    {
        // The same INDEX consumer over the shapes that broadcast COMPLETELY: (2,3) of A1:C3*E1:E3 is 6*2,
        // (3,1) is 7*3, and (3,2) of the outer product is 30*20. Today #VALUE!.
        await Assert.That(Num(OnGrid("=INDEX(A1:C3*E1:E3,2,3)"))).IsEqualTo(12.0);
        await Assert.That(Num(OnGrid("=INDEX(A1:C3*E1:E3,3,1)"))).IsEqualTo(21.0);
        await Assert.That(Num(OnGrid("=INDEX(E1:E3*E5:G5,3,2)"))).IsEqualTo(60.0);
    }

    // --- Composition: chained operators, a comparison, a different operator ---

    [Test]
    public async Task ChainedAndMixedOperators_BroadcastAtEveryLevel()
    {
        // A 3x3 against a 3x1 against a 1x3: the two LEAVES project at every level, but the intermediate
        // A1:C3*E1:E3 is itself a 3x3, so the outer extent is the composite's OWN and its projection is the
        // identity — this pin does not read a composite at a foreign extent (that is
        // ACompositeOperand_ReadAtAForeignExtent_ProjectsBeforeAskingItsChildren). Today #VALUE! for all three.
        await Assert.That(Num(OnGrid("=SUM(A1:C3*E1:E3*E5:G5)"))).IsEqualTo(2280.0);

        // The rule is the OPERATOR's, not multiplication's: a comparison broadcasts (2 hits in row 2 times 2,
        // 3 in row 3 times 3) and so does addition (45 + 18).
        await Assert.That(Num(OnGrid("=SUM((A1:C3>4)*E1:E3)"))).IsEqualTo(13.0);
        await Assert.That(Num(OnGrid("=SUM(A1:C3+E1:E3)"))).IsEqualTo(63.0);
    }

    [Test]
    public async Task NestedExtents_ProjectThroughEveryLevel()
    {
        // (A1:C3*H1:H2)*E5:G5: the inner product is a 3x3 with an uncovered row 3 and the outer operand is a
        // 1x3 row, so the #N/A survives the second fold and SIX positions stay countable. What this pin does
        // NOT prove: that the composite projects — A1:C3 dominates every fold, so the inner 3x3 is read at its
        // own 3x3 extent and its projection is the identity (Task 2's reviewer, 2026-09-10). The pin that
        // reads a composite at a DIFFERENT extent is
        // ACompositeOperand_ReadAtAForeignExtent_ProjectsBeforeAskingItsChildren. Today #VALUE! / 0.
        await Assert.That(OnGrid("=SUM(A1:C3*H1:H2*E5:G5)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(A1:C3*H1:H2*E5:G5)"))).IsEqualTo(6.0);
    }

    // --- Composites read at a FOREIGN extent (Task 3 witnesses) ---
    //
    // Every value below: Aspose.Cells 26.6.0, measured 2026-09-10, CSE column (plain entry answers #VALUE! for
    // every SUM and 0 for every COUNT, and agrees on every INDEX). "Today" is this tree at 2c8ff98, where the
    // four composites still carried the equal-shape guard.

    [Test]
    public async Task ACompositeOperand_ReadAtAForeignExtent_ProjectsBeforeAskingItsChildren()
    {
        // The controller's fixture: A1:A3*H1:H2 is a 3x1 COMPOSITE ([1, 8, #N/A]) inside a 3x3 whose other
        // operand is a 1x3 row. The composite is asked at 3x3, an extent that is not its own, so it must
        // project the parent's index onto its 3x1 BEFORE handing its own index to its two leaves — the
        // pre-Phase-10 guard refused the extent instead. Rows 1-2 are covered (six numbers), row 3 is #N/A.
        // Today #VALUE! / 0 / #VALUE! ×3.
        await Assert.That(OnGrid("=SUM((A1:A3*H1:H2)*E5:G5)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT((A1:A3*H1:H2)*E5:G5)"))).IsEqualTo(6.0);
        await Assert.That(Num(OnGrid("=INDEX((A1:A3*H1:H2)*E5:G5,1,1)"))).IsEqualTo(10.0);
        await Assert.That(Num(OnGrid("=INDEX((A1:A3*H1:H2)*E5:G5,2,3)"))).IsEqualTo(240.0);
        await Assert
            .That(OnGrid("=INDEX((A1:A3*H1:H2)*E5:G5,3,1)"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // The same composite on the RIGHT of the outer operator: the side does not matter to the projection.
        // Today #VALUE! / 0.
        await Assert.That(OnGrid("=SUM(E5:G5*(A1:A3*H1:H2))")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(E5:G5*(A1:A3*H1:H2))"))).IsEqualTo(6.0);

        // Composites that broadcast COMPLETELY at the foreign extent — a 3x1 column composite and a 1x3 row
        // composite, each read at 3x3 with nothing uncovered: (2+3+4)*60 and (11+21+31)*6. Today #VALUE! ×3.
        await Assert.That(Num(OnGrid("=SUM((E1:E3+1)*E5:G5)"))).IsEqualTo(540.0);
        await Assert.That(Num(OnGrid("=SUM((E5:G5+1)*E1:E3)"))).IsEqualTo(378.0);
        await Assert.That(Num(OnGrid("=INDEX((E1:E3+1)*E5:G5,3,2)"))).IsEqualTo(80.0);
    }

    [Test]
    public async Task AnIfComposite_ProjectsAtAForeignExtent_AndReadsItsOwnChildrenAtItsOwn()
    {
        // An IF as the composite read at a foreign extent: IF(E1:E3>1,1,0) is a 3x1 [0, 1, 1] inside a 3x3
        // (two true rows × 60 = 120, nine numbers), and IF(H1:H2>1,1,0) is a 2x1 inside a 3x3 with row 3
        // uncovered. Today #VALUE! / 0 / #VALUE! / 0.
        await Assert.That(Num(OnGrid("=SUM(IF(E1:E3>1,1,0)*E5:G5)"))).IsEqualTo(120.0);
        await Assert.That(Num(OnGrid("=COUNT(IF(E1:E3>1,1,0)*E5:G5)"))).IsEqualTo(9.0);
        await Assert.That(OnGrid("=SUM(IF(H1:H2>1,1,0)*A1:C3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(IF(H1:H2>1,1,0)*A1:C3)"))).IsEqualTo(6.0);

        // The other direction: an IF whose CONDITION is a composite (H1:H2*1, a 2x1) narrower than the IF's
        // own 3x3 extent. The IF must ask its condition at 3x3 so that row 3 is #N/A and rows 1-2 read the
        // real comparison. Today 4 / 2 — silent wrong NUMBERS: the IF took its condition's 2x1 as its extent
        // and read A1:C3 at 2x1, so only A2 was ever summed.
        await Assert
            .That(OnGrid("=SUM(IF((H1:H2*1)>1,A1:C3,0))"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(IF((H1:H2*1)>1,A1:C3,0))"))).IsEqualTo(6.0);

        // And an IF whose taken BRANCH is a composite (E1:E3*2, a 3x1) read at the IF's 3x3 extent from a
        // 1x3 condition: columns 2-3 are true and each reads the whole doubled column, 2*(2+4+6). Today
        // #VALUE! / 1 — the branch composite refused the 1x3 extent, and the one false 0 was counted.
        await Assert.That(Num(OnGrid("=SUM(IF(E5:G5>10,E1:E3*2,0))"))).IsEqualTo(24.0);
        await Assert.That(Num(OnGrid("=COUNT(IF(E5:G5>10,E1:E3*2,0))"))).IsEqualTo(9.0);
    }

    [Test]
    public async Task AUnaryAndALiftedComposite_ProjectAtAForeignExtent()
    {
        // Phase 8's two composites read at an extent that is not their own. -H1:H2 is a 2x1 UnaryOperand
        // inside a 3x3 (row 3 uncovered); ABS(H1:H2) is a 2x1 LiftedFunctionOperand inside the same 3x3, and
        // inside a 2x3 against the row vector where nothing is uncovered (3*60); ROUND(E1:E3,0) is a 3x1
        // lifted operand read at 3x3 (6*60). Today #VALUE! / 0 / #VALUE! / 0 / #VALUE! / #VALUE!.
        await Assert.That(OnGrid("=SUM(-H1:H2*A1:C3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(-H1:H2*A1:C3)"))).IsEqualTo(6.0);
        await Assert.That(OnGrid("=SUM(ABS(H1:H2)*A1:C3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(ABS(H1:H2)*A1:C3)"))).IsEqualTo(6.0);
        await Assert.That(Num(OnGrid("=SUM(ABS(H1:H2)*E5:G5)"))).IsEqualTo(180.0);
        await Assert.That(Num(OnGrid("=SUM(ROUND(E1:E3,0)*E5:G5)"))).IsEqualTo(360.0);

        // Three levels: a lifted composite over a binary composite, read at a wider extent still. ROUND
        // projects the 3x3 index onto its 3x1, then hands its own index to A1:A3*H1:H2, which projects again
        // onto its leaves; the #N/A at row 3 survives both. Today #VALUE! / 0.
        await Assert
            .That(OnGrid("=SUM(ROUND(A1:A3*H1:H2,0)*E5:G5)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(ROUND(A1:A3*H1:H2,0)*E5:G5)"))).IsEqualTo(6.0);
    }

    // --- IF: the condition AND both branches fold into one extent ---

    [Test]
    public async Task If_BroadcastsTheConditionAgainstTheBranch()
    {
        // A 3x1 condition selecting whole ROWS of a 3x3 branch (rows 2 and 3 = 15 + 24) and a 1x3 condition
        // selecting whole COLUMNS (columns B and C = 15 + 18). Today #VALUE! for both.
        await Assert.That(Num(OnGrid("=SUM(IF(E1:E3>1,A1:C3,0))"))).IsEqualTo(39.0);
        await Assert.That(Num(OnGrid("=SUM(IF(E5:G5>10,A1:C3,0))"))).IsEqualTo(33.0);
    }

    [Test]
    public async Task If_FoldsTheBranchShapeIntoItsOwnExtent()
    {
        // A 3x1 condition against a 1x3 BRANCH is a 3x3 result — 60 per true row, rows 2 and 3 = 120. This
        // is impossible while the IF's extent is the condition's alone (it would be 3x1). Today #VALUE!.
        await Assert.That(Num(OnGrid("=SUM(IF(E1:E3>1,E5:G5,0))"))).IsEqualTo(120.0);

        // The same fold with a 2x1 condition and a 1x3 branch: a 2x3 result, one true row = 60. Today
        // #VALUE!.
        await Assert.That(Num(OnGrid("=SUM(IF(H1:H2>1,E5:G5,0))"))).IsEqualTo(60.0);
    }

    [Test]
    public async Task If_OverAnUncoveredCondition_MarksTheTail_InsteadOfCountingElseValues()
    {
        // The phase's reason to exist in the IF arm, and the only pins here whose TODAY value is a NUMBER
        // rather than an error — a wrong answer that no error surfaced. Today IfOperand takes the
        // CONDITION's extent, so a 2x1 condition over a 3x3 branch reads the branch at a mismatched extent
        // and quietly counts the 0 else-values as numbers: COUNT(IF(H1:H2>1,A1:C3,0)) is 1 today (6 on the
        // oracle), COUNT(IF(H1:H2>1,A1:C3)) is 0 today (3), and COUNT(IF(A1:C3>4,H1:H2,0)) is 4 today (6).
        //
        // Under the rule the extent is 3x3: row 1 is false (three 0s), row 2 true (4,5,6), row 3 uncovered
        // by the condition (#N/A x3) — so SUM is #N/A while COUNT is 6.
        await Assert.That(OnGrid("=SUM(IF(H1:H2>1,A1:C3,0))")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(IF(H1:H2>1,A1:C3,0))"))).IsEqualTo(6.0);

        // Without an else branch the false row yields logical FALSE, which COUNT does not count, so only
        // row 2's three numbers remain countable.
        await Assert.That(Num(OnGrid("=COUNT(IF(H1:H2>1,A1:C3))"))).IsEqualTo(3.0);

        // The mirror image: a 3x3 condition over a 2x1 BRANCH. Row 3's condition is true everywhere but the
        // branch does not cover it, so those three positions are #N/A; the six false/covered ones are
        // numbers (four 0s and two 2s).
        await Assert.That(OnGrid("=SUM(IF(A1:C3>4,H1:H2,0))")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(IF(A1:C3>4,H1:H2,0))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task If_WithTwoVectorBranches_FoldsAllThreeShapes()
    {
        // Verifier correction m6. Both branches are vectors of DIFFERENT shape, so the extent comes from all
        // three operands: a 3x1 condition, a 3x3 true branch and a 1x3 false branch fold to 3x3 — row 1
        // takes E5:G5 (60) and rows 2-3 take A1:C3 (39). Today #VALUE! for the SUM and #REF! for the INDEX —
        // today's extent is the condition's 3x1 alone, so column 2 is out of bounds.
        await Assert.That(Num(OnGrid("=SUM(IF(E1:E3>1,A1:C3,E5:G5))"))).IsEqualTo(99.0);
        await Assert.That(Num(OnGrid("=INDEX(IF(E1:E3>1,A1:C3,E5:G5),1,2)"))).IsEqualTo(20.0);

        // A 2x1 FALSE branch that the fold widens to 3 rows: row 1 is false and reads H1 three times (3),
        // rows 2-3 take the 1x3 row branch (120). Nothing is uncovered, because the row taken from the short
        // branch is row 1. Today #VALUE! / 0.
        await Assert.That(Num(OnGrid("=SUM(IF(E1:E3>1,E5:G5,H1:H2))"))).IsEqualTo(123.0);
        await Assert.That(Num(OnGrid("=COUNT(IF(E1:E3>1,E5:G5,H1:H2))"))).IsEqualTo(9.0);

        // The same fixture with the SHORT branch taken on the uncovered row: a 3x1 condition and a 2x1 true
        // branch fold to 3x1 (not 3x3 — both are columns), row 1 is false (0), row 2 reads H2 (2) and row 3
        // is uncovered. Today #VALUE! for the SUM and 1 for the COUNT — silently one number where the oracle
        // has two. The INDEX line is a REGRESSION pin, not a RED one: element (1,1) is already 0 today
        // (row 1 takes the else branch either way) and must stay 0 once the extent widens.
        await Assert.That(OnGrid("=SUM(IF(E1:E3>1,H1:H2,0))")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(IF(E1:E3>1,H1:H2,0))"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=INDEX(IF(E1:E3>1,H1:H2,0),1,1)"))).IsEqualTo(0.0);
    }

    // --- Error precedence and blanks ---

    [Test]
    public async Task AtAnUncoveredPosition_TheOtherOperandsOwnErrorStillWins()
    {
        // The one place where a "correct" resolver could still be wrong. If the projection were applied to
        // the whole OPERATION — refusing the position before reading either side — (3,1) would be #N/A both
        // ways. It is not: with A3 = #DIV/0! the operands are read left then right and the FIRST error found
        // is the answer, so the same position gives #DIV/0! or #N/A depending on which operand comes first.
        //
        // Today the first line already answers #DIV/0! — a REGRESSION pin: the left operand's error already
        // beats the right operand's own marker, which today is #VALUE! rather than #N/A. The other four are
        // #VALUE! today.
        await Assert
            .That(OnGridWithErrorInA3("=INDEX(A1:C3*H1:H2,3,1)"))
            .IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(OnGridWithErrorInA3("=INDEX(H1:H2*A1:C3,3,1)"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // (3,2) has no error on the left — B3 is 8 — so the uncovered right operand answers there.
        await Assert
            .That(OnGridWithErrorInA3("=INDEX(A1:C3*H1:H2,3,2)"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // The streamed consumer reports the first error in row-major scan order, which is (3,1) in both.
        await Assert.That(OnGridWithErrorInA3("=SUM(A1:C3*H1:H2)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(OnGridWithErrorInA3("=SUM(H1:H2*A1:C3)"))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task AnErrorEarlierInScanOrder_BeatsALaterUncoveredPosition()
    {
        // B2 = #DIV/0! sits at (2,2) and the uncovered tail starts at (3,1), so the error wins on order
        // alone — the broadcast changes WHICH positions exist, not the "first error in scan order" rule.
        // Today the two SUMs are #VALUE! while the INDEX is already #DIV/0! (a regression pin, for the same
        // reason as the test above).
        await Assert.That(OnGridWithErrorInB2("=SUM(A1:C3*E1:E3)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(OnGridWithErrorInB2("=SUM(A1:C3*H1:H2)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert
            .That(OnGridWithErrorInB2("=INDEX(A1:C3*H1:H2,2,2)"))
            .IsEqualTo(ErrorValue.DivByZero);

        // E2 is BLANK in this fixture: a blank element broadcasts as 0, it does not refuse. Row 2 therefore
        // contributes 4*0, #DIV/0! and 6*0 — eight of the nine elements are numbers (0 today). SUM(E1:E3*3)
        // is already 12 today (an equal-shape operand against a scalar never needed broadcasting); it is
        // here as the companion that shows the blank is the reason COUNT is 8 and not 9.
        await Assert.That(Num(OnGridWithErrorInB2("=COUNT(A1:C3*E1:E3)"))).IsEqualTo(8.0);
        await Assert.That(Num(OnGridWithErrorInB2("=SUM(E1:E3*3)"))).IsEqualTo(12.0);
    }

    // --- The other consumers over a broadcast array ---

    [Test]
    public async Task MaxMinAndAverage_OverAMismatchedPair_AreNotAvailable()
    {
        // Verifier correction m6. The MIN/MAX/AVERAGE family folds the same stream as SUM, so an uncovered
        // position propagates there too. Today all five are #VALUE!.
        await Assert.That(OnGrid("=MAX(A1:C3*H1:H2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=MIN(A1:C3*H1:H2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=AVERAGE(A1:C3*H1:H2)")).IsEqualTo(ErrorValue.NotAvailable);

        // The companions on a shape that broadcasts completely, so the family is pinned on both sides of the
        // rule: max of 1,2,3,8,10,12,21,24,27 and 108/9.
        await Assert.That(Num(OnGrid("=MAX(A1:C3*E1:E3)"))).IsEqualTo(27.0);
        await Assert.That(Num(OnGrid("=AVERAGE(A1:C3*E1:E3)"))).IsEqualTo(12.0);
    }

    [Test]
    public async Task SumProduct_ConsumesABroadcastArgument_ButNeverBroadcastsItsOwn()
    {
        // Verifier correction m6. SUMPRODUCT keeps a dimension rule of its OWN, stricter than the operator's:
        // its ARGUMENTS must match exactly (a 1x1 included — see MathAggregateTests for that list), while a
        // broadcast expression INSIDE one argument is consumed as the broadcast array.
        //
        // A 3x3 comparison broadcast against a 3x1: 2 + 3 + 3 hits. Today #VALUE!.
        await Assert.That(Num(OnGrid("=SUMPRODUCT(--(A1:C3>E1:E3))"))).IsEqualTo(8.0);

        // The same idiom over the uncovered pair: the #N/A elements propagate through SUMPRODUCT's fold.
        // Today #VALUE!.
        await Assert
            .That(OnGrid("=SUMPRODUCT(--(A1:C3>H1:H2))"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // And SUMPRODUCT's own rule is untouched by this phase: a COMPUTED 3x3 against a bare 3x1 argument
        // is still its own mismatch, #VALUE! — not a broadcast. Already #VALUE! today, so this is a
        // regression pin: it must NOT become 13 when the operator learns to broadcast.
        await Assert.That(OnGrid("=SUMPRODUCT((A1:C3>4)*1,E1:E3)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task TwoRowsOfDifferentLength_FollowTheSameRuleAsTwoColumns()
    {
        // Verifier correction m6. The rule is per-axis and symmetric: a 1x2 against a 1x3 is a 1x3 whose
        // third position is uncovered — 100, 400, #N/A. Today #VALUE! / 0 / #VALUE!.
        await Assert.That(OnGrid("=SUM(E5:F5*E5:G5)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(E5:F5*E5:G5)"))).IsEqualTo(2.0);
        await Assert.That(OnGrid("=INDEX(E5:F5*E5:G5,1,3)")).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Unary_OverABroadcastArray_KeepsTheUncoveredTail()
    {
        // Verifier correction m6. The UnaryOperand path (Phase 8) projects through the same resolver, so
        // negating a broadcast array neither hides nor multiplies the #N/A elements. Today #VALUE! / 0.
        await Assert.That(OnGrid("=SUM(-(A1:C3*H1:H2))")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(-(A1:C3*H1:H2))"))).IsEqualTo(6.0);
    }

    // --- Phase 8's lifted operands over a broadcast (these need item 8, not only the operator) ---

    [Test]
    public async Task ALiftedFunctionsArguments_BroadcastLikeAnOperatorsOperands()
    {
        // ROUND is an ArrayLifting.Elementwise built-in, so its argument shapes fold by the same rule and
        // each argument is read through the same projection. Today #VALUE! / #VALUE! / 0 / #VALUE! / 0.
        await Assert.That(Num(OnGrid("=SUM(ROUND(A1:C3,E1:E3))"))).IsEqualTo(45.0);

        // A 3x3 first argument against a 2x1 second: row 3 uncovered, six covered.
        await Assert.That(OnGrid("=SUM(ROUND(A1:C3,H1:H2))")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(ROUND(A1:C3,H1:H2))"))).IsEqualTo(6.0);

        // A 3x1 against a 1x3 in a lifted function's slots is the same outer product as in an operator: nine
        // elements, each ROUND(E[r], G[c]) = E[r], so 3*(1+2+3) = 18. A per-axis Math.Max fold without the
        // projection would build the same 3x3 extent and then fail every element.
        await Assert.That(Num(OnGrid("=SUM(ROUND(E1:E3,E5:G5))"))).IsEqualTo(18.0);

        // A 2x2 against a 3x1 → 3x2, four covered positions and an uncovered row 3.
        await Assert.That(Num(OnGrid("=COUNT(ROUND(A1:B2,E1:E3))"))).IsEqualTo(4.0);
    }

    [Test]
    public async Task AnErrorConsumingLiftedBody_SeesARealNotAvailablePerElement()
    {
        // The user-visible proof that the uncovered tail is a per-ELEMENT #N/A and not a whole-array
        // refusal: IFERROR recovers exactly the three uncovered positions (36 = 1+2+3+8+10+12, and 33 when
        // each is replaced by -1), and ISNA marks exactly three of the nine.
        //
        // Today: SUM(IFERROR(…,0)) = 0 and SUM(IFERROR(…,-1)) = -9 — IFERROR is lifted at HEAD and consumes
        // the per-element #VALUE! marker nine times, so the numbers are wrong in a way no error reveals
        // (verifier correction M3). SUM(ISNA(…)*1) = 0 today, because the marker is #VALUE!, not #N/A.
        await Assert.That(Num(OnGrid("=SUM(IFERROR(A1:C3*H1:H2,0))"))).IsEqualTo(36.0);
        await Assert.That(Num(OnGrid("=SUM(IFERROR(A1:C3*H1:H2,-1))"))).IsEqualTo(33.0);
        await Assert.That(Num(OnGrid("=SUM(ISNA(A1:C3*H1:H2)*1)"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task AUnaryAndAConcatenation_BroadcastThroughTheirLiftedOperands()
    {
        // -E1:E3 is a lifted unary over a 3x1 and E5:G5 a 1x3, so the negated outer product is -360; and
        // E1:E3&E5:G5 concatenates 1&10, 1&20, … — nine two-and-three-character strings, LEN summing to 27.
        // Today #VALUE! for both.
        await Assert.That(Num(OnGrid("=SUM(-E1:E3*E5:G5)"))).IsEqualTo(-360.0);
        await Assert.That(Num(OnGrid("=SUM(LEN(E1:E3&E5:G5))"))).IsEqualTo(27.0);
    }

    // --- ROW/COLUMN of a rectangle are VECTORS, not rectangles (verifier correction B1) ---

    [Test]
    public async Task RowAndColumn_OfARectangle_AreVectorsThatBroadcast()
    {
        // Verifier correction B1, which the phase file makes BLOCKING: this engine fabricates an MxN
        // rectangle for ROW/COLUMN of a rectangle ("every cell in row r shares the row number"), while the
        // oracle produces an Nx1 COLUMN for ROW and a 1xM ROW for COLUMN. The MxN fabrication was only ever
        // needed because the operand tree could not broadcast; with broadcasting the rectangle idioms still
        // work, and the element COUNT is no longer nine.
        //
        // Today: #VALUE! (the fabricated 3x3 against a 3x1 is a shape mismatch) / 18 / 9 / #VALUE!.
        await Assert.That(Num(OnGrid("=SUM(ROW(A1:C3)*E1:E3)"))).IsEqualTo(14.0);
        await Assert.That(Num(OnGrid("=SUM(ROW(A1:C3))"))).IsEqualTo(6.0);
        await Assert.That(Num(OnGrid("=COUNT(ROW(A1:C3))"))).IsEqualTo(3.0);

        // The outer product of the two position vectors — a shape that cannot exist while both are
        // rectangles of the SAME range's dimensions.
        await Assert.That(Num(OnGrid("=SUM(ROW(A1:A3)*COLUMN(A1:C1))"))).IsEqualTo(36.0);
    }

    [Test]
    public async Task RowAndColumnVectors_KeepTheRectangleIdioms_AndLoseTheRectanglesElementCount()
    {
        // The consequences of B1 a consumer can observe (Task 2 pins; Aspose.Cells 26.6.0, measured
        // 2026-09-10, CSE column — plain entry answers 1 for the bare COUNT and #VALUE! for the products).
        // Measured on this tree at 2f77c97 (the MxN rectangles): the three idiom sums were ALREADY 108, 96
        // and 228 — regression pins, the rectangle and the vector agree there — while the element-count
        // lines answered 9 / 1 / 2 / #VALUE! / 0.
        //
        // The idioms: the 3x1 row vector times the rectangle weights each row by its number (6 + 30 + 72),
        // the 1x3 column vector weights each column (12 + 30 + 54), and both at once (row * column * cell).
        await Assert.That(Num(OnGrid("=SUM(ROW(A1:C3)*A1:C3)"))).IsEqualTo(108.0);
        await Assert.That(Num(OnGrid("=SUM(COLUMN(A1:C3)*A1:C3)"))).IsEqualTo(96.0);
        await Assert.That(Num(OnGrid("=SUM(ROW(A1:C3)*COLUMN(A1:C3)*A1:C3)"))).IsEqualTo(228.0);

        // The element count is the vector's: COLUMN(A1:C3) has three elements, ROW(A1:C3) has no column 3
        // for INDEX and no 4th smallest for SMALL, and against the 2x1 H1:H2 it is a 3x1 with row 3
        // uncovered — two covered elements, not six.
        await Assert.That(Num(OnGrid("=COUNT(COLUMN(A1:C3))"))).IsEqualTo(3.0);
        await Assert.That(OnGrid("=INDEX(ROW(A1:C3),1,3)")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(OnGrid("=SMALL(ROW(A1:C3),4)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(OnGrid("=SUM(ROW(A1:C3)*H1:H2)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(ROW(A1:C3)*H1:H2)"))).IsEqualTo(2.0);
    }

    // --- The COLUMN-axis uncovered path (Task 2 witnesses) ---

    [Test]
    public async Task AnUncoveredColumn_IsMarkedExactlyLikeAnUncoveredRow()
    {
        // Before these pins SUM(E5:F5*E5:G5) above was the only streamed witness of the column-axis half
        // of Broadcasting.TryProject, and every worked example in the design is row-axis. Aspose.Cells
        // 26.6.0, measured 2026-09-10, CSE column (plain entry answers #VALUE! / 0 for the SUMs and COUNTs
        // and agrees on every INDEX). Measured on this tree at 2f77c97: #VALUE! for every SUM and INDEX
        // below, 0 for every COUNT — the equal-shape guard's per-element fill.
        //
        // 3x2 against 3x3 → 3x3: the LEFT operand has no column 3, so (1,3), (2,3) and (3,3) are #N/A and
        // the six positions in columns 1-2 are the squares 1, 4, 16, 25, 49, 64 — an uncovered COLUMN, not a
        // tail of the row-major index range, which is what distinguishes this path from the row-axis pins.
        await Assert.That(OnGrid("=SUM(A1:B3*A1:C3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(A1:B3*A1:C3)"))).IsEqualTo(6.0);
        await Assert.That(OnGrid("=INDEX(A1:B3*A1:C3,1,3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=INDEX(A1:B3*A1:C3,2,3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnGrid("=INDEX(A1:B3*A1:C3,3,3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=INDEX(A1:B3*A1:C3,1,1)"))).IsEqualTo(1.0);
        await Assert.That(Num(OnGrid("=INDEX(A1:B3*A1:C3,3,2)"))).IsEqualTo(64.0);

        // A 1x2 ROW against the 3x3: the row repeats down every row AND leaves column 3 uncovered — both
        // clauses on one operand. (3,2) is 20*8; (3,3) is #N/A.
        await Assert.That(OnGrid("=SUM(E5:F5*A1:C3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(E5:F5*A1:C3)"))).IsEqualTo(6.0);
        await Assert.That(Num(OnGrid("=INDEX(E5:F5*A1:C3,3,2)"))).IsEqualTo(160.0);
        await Assert.That(OnGrid("=INDEX(E5:F5*A1:C3,3,3)")).IsEqualTo(ErrorValue.NotAvailable);

        // The covered element of the 1x2-against-1x3 pin above: 20*20.
        await Assert.That(Num(OnGrid("=INDEX(E5:F5*E5:G5,1,2)"))).IsEqualTo(400.0);

        // The same path through a POSITION operand: COLUMN(A1:B1) is the 1x2 row [1,2] (B1), so against
        // the 3x3 it is #N/A down column 3 and column c's number times the cell elsewhere — (2,2) is 2*5.
        await Assert.That(OnGrid("=SUM(COLUMN(A1:B1)*A1:C3)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=COUNT(COLUMN(A1:B1)*A1:C3)"))).IsEqualTo(6.0);
        await Assert
            .That(OnGrid("=INDEX(COLUMN(A1:B1)*A1:C3,2,3)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnGrid("=INDEX(COLUMN(A1:B1)*A1:C3,2,2)"))).IsEqualTo(10.0);
    }
}
