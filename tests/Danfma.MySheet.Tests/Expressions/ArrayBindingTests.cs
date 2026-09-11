using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Tests.Expressions.SelectionProducerFixture;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 11c (array bindings) — the acceptance pins for a computed array that crosses a BINDING site: a
/// <c>LET</c> binding, the chosen branch of <c>CHOOSE</c>, the operand of unary <c>+</c>, and a workbook defined
/// name whose definition is a computed-array expression. <b>Every non-guard test here is RED at the commit
/// that adds it</b>: <c>NamedReferences.CaptureValue</c>'s fall-through evaluates a non-reference binding as an
/// ordinary scalar, so a producer collapses to its top-left silently (5 where the answer is 14, 1 where it is
/// 2) and an operator over a range or an array <c>IF</c> collapses to a loud <c>#VALUE!</c>. Each row's comment
/// names today's value beside the oracle's, so a row that is green for the wrong reason is visible.
/// </summary>
/// <remarks>
/// <para>
/// <b>Golden values.</b> The documented behaviour comes from three official Microsoft pages on
/// support.microsoft.com, fetched <b>2026-09-11</b>: "LET function" (34842dd8-b92b-4d3f-b325-b8b8f9908999) —
/// "LET allows you to call the expression by name and for Excel to calculate it once"; "CHOOSE function"
/// (fc5c184f-cb62-4ec7-a46e-38653b98f5bc) — "If index_num is less than 1 or greater than the number of the
/// last value in the list, CHOOSE returns the #VALUE! error value" and "The value arguments to CHOOSE can be
/// range references as well as single values"; and "UNIQUE function" (c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e) —
/// "The UNIQUE function will return an array, which will spill if it's the final result of a formula", the
/// sentence that names a producer's result an ARRAY, which is what a binding must carry.
/// </para>
/// <para>
/// <b>Where a page is silent or disagrees, the oracle decides</b> (project rule P0, addendum 2026-09-09):
/// "Excel" means <b>Aspose.Cells 26.6.0 as measured</b>. Every number below was measured on that version on
/// <b>2026-09-11</b> against this file's fixture, in BOTH entry modes — <c>plain</c> (<c>Cell.Formula</c>) and
/// <c>CSE</c> (<c>Cell.SetArrayFormula(f, 1, 1)</c>) — one formula per workbook, in cell <c>H20</c>. The two
/// columns agree for every row except the seven whose test names the column in its own name
/// (<c>SUM(+(A1:A3*2))</c>, <c>SUM(+IF(A1:A3>0,A1:A3))</c>, <c>SUM(-(+A1:A3))</c> and the four Phase 11a
/// <c>LET(r,A1:A3*1,…)</c> criteria rows), and where they split <b>the CSE column is the target</b>, because
/// the mini-CSE implements the array-entered rule everywhere. Modes are never mixed inside one assertion.
/// </para>
/// <para>
/// Fixture: <see cref="SelectionProducerFixture.Grid"/> (<c>A1:A3</c> = 5, 0, 9, so
/// <c>FILTER(A1:A3,A1:A3>0)</c> is the two-row 5, 9 with sum 14) plus two workbook names: <c>ProdName</c> is
/// <c>FILTER(Sheet1!$A$1:$A$3,Sheet1!$A$1:$A$3>0)</c> and <c>OpName</c> is <c>Sheet1!$A$1:$A$3*2</c>.
/// </para>
/// </remarks>
public class ArrayBindingTests
{
    private static (Workbook Workbook, Sheet Sheet) Named()
    {
        var (workbook, sheet) = Grid();
        workbook.DefineName("ProdName", "FILTER(Sheet1!$A$1:$A$3,Sheet1!$A$1:$A$3>0)");
        workbook.DefineName("OpName", "Sheet1!$A$1:$A$3*2");

        return (workbook, sheet);
    }

    private static object? Calc(string formula)
    {
        var (workbook, sheet) = Named();

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // The CELL path: the formula is stored in H20 and read back through Workbook.GetCellValue, the only path
    // that crosses the cell boundary (Workbook.cs, "CAPTURE, not plain Evaluate"). Row 20 lies outside A1:A3,
    // so an implicit intersection there would answer 0, never 5 — a 5 through this path is the top-left rule.
    private static object? InCell(string formula)
    {
        var (workbook, sheet) = Named();
        sheet["H20"] = ExpressionParser.Parse(formula, sheet);

        return workbook.GetCellValue("Sheet1", "H20").AsObject();
    }

    // The oracle's cell text as the value the engine must answer: "#REF!" is the error, "14" the number.
    // ErrorValue's equality is on the code string, so a code the engine has no Error member for still pins.
    private static object Oracle(string text) =>
        text.StartsWith('#') ? new ErrorValue(text) : double.Parse(text);

    // ================================================================================================
    // LET: an array binding streams whole to every consumer of the name
    // ================================================================================================

    [Test]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),SUM(f))", "14")] // today 5 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),ROWS(f))", "2")] // today 1 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),INDEX(f,2))", "9")] // today #REF! — loud
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),SUM(f)+ROWS(f))", "16")] // today 6 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),SUM(f*2))", "28")] // today 10 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),SUM(LEN(f)))", "2")] // today 1 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),COUNTA(f))", "2")] // today 1 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),SUMPRODUCT(f))", "14")] // today 5 — silent
    [Arguments("=LET(a,FILTER(A1:A3,A1:A3>0),b,SORT(a),SUM(b))", "14")] // today 5 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),g,f*2,SUM(g))", "28")] // today 10 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),LET(g,SORT(f),SUM(g)))", "14")] // today 5 — silent
    [Arguments("=LET(x,A1:A3*2,SUM(x))", "28")] // today #VALUE! — loud
    [Arguments("=LET(x,IF(A1:A3>0,A1:A3),SUM(x))", "14")] // today #VALUE! — loud
    [Arguments("=SUM(LET(f,FILTER(A1:A3,A1:A3>0),f))", "14")] // today 5 — silent
    public async Task ALetBoundArray_StreamsWholeToItsConsumer(string formula, string oracle)
    {
        // Oracle 26.6.0, 2026-09-11: the value named, in BOTH entry modes, for every row. The two loud rows
        // are the binding shapes CaptureValue cannot even collapse — an operator over a range and an array IF
        // evaluate to #VALUE! on the scalar path — and the silent rows are a producer's FirstElement rule
        // applied one site too early: SUM([5]) = 5, ROWS([5]) = 1, SUM([5]*2) = 10.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    public async Task ALetBoundEmptyProducer_KeepsItsCalcError()
    {
        // Oracle 26.6.0, 2026-09-11: #CALC! in BOTH modes — an empty FILTER bound by LET is still the 1x1
        // #CALC! singleton when SUM reads it, so the binding must carry the error, not swallow it into a
        // blank or a 0. GUARD: green today, because the collapse is invisible on a 1x1 — the FirstElement of
        // the #CALC! singleton is #CALC! itself — and it must stay green when the binding carries the operand
        // instead: an empty operand bound and then summed is still #CALC!, never 0.
        await Assert
            .That(Calc("=LET(f,FILTER(A1:A3,A1:A3>100),SUM(f))"))
            .IsEqualTo(ErrorValue.Calculation);
    }

    // ================================================================================================
    // CHOOSE: the chosen branch streams whole; index errors are the loud 1x1 they already are
    // ================================================================================================

    [Test]
    [Arguments("=SUM(CHOOSE(1,FILTER(A1:A3,A1:A3>0)))", "14")] // today 5 — silent
    [Arguments("=SUM(CHOOSE(2,0,FILTER(A1:A3,A1:A3>0)))", "14")] // today 5 — silent
    [Arguments("=ROWS(CHOOSE(1,FILTER(A1:A3,A1:A3>0)))", "2")] // today 1 — silent
    public async Task AChosenArray_StreamsWholeToItsConsumer(string formula, string oracle)
    {
        // Oracle 26.6.0, 2026-09-11: 14 / 14 / 2 in BOTH modes. Today the chosen branch goes through
        // LookupFunctions' CaptureValue and collapses to its top-left before SUM or ROWS can see it.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    [Arguments("=CHOOSE(1/0,FILTER(A1:A3,A1:A3>0))", "#DIV/0!")]
    [Arguments("=SUM(CHOOSE(1/0,FILTER(A1:A3,A1:A3>0)))", "#DIV/0!")]
    [Arguments("=CHOOSE(3,FILTER(A1:A3,A1:A3>0))", "#VALUE!")]
    [Arguments("=SUM(CHOOSE(3,FILTER(A1:A3,A1:A3>0)))", "#VALUE!")]
    public async Task AChooseIndexError_IsTheLoudScalar_EvenWithAnArrayBranch(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11: #DIV/0! for an error index and #VALUE! for an out-of-range one, in BOTH
        // modes, bare and under SUM alike — the CHOOSE page's rule, and the array branch changes nothing
        // because it is never chosen. GUARD: green today on the scalar path (the index is evaluated before
        // any branch is); whatever arm Phase 3 adds must keep answering the index error as a loud 1x1 and
        // never probe or build the branch first.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    // ================================================================================================
    // Unary +: transparent to an array, still a reference over a bare range
    // ================================================================================================

    [Test]
    [Arguments("=SUM(+FILTER(A1:A3,A1:A3>0))", "14")] // today 5 — silent
    [Arguments("=ROWS(+FILTER(A1:A3,A1:A3>0))", "2")] // today 1 — silent
    [Arguments("=SUM(-(+FILTER(A1:A3,A1:A3>0)))", "-14")] // today -5 — silent
    public async Task AUnaryPlusOverAProducer_StreamsWholeToItsConsumer(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11: 14 / 2 / -14 in BOTH modes. Today UnaryOperation.Evaluate captures the
        // operand as a value, which for a producer is its FirstElement, and the '+' is opaque to the probe.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    [Arguments("=SUM(+(A1:A3*2))", "28")] // today #VALUE! — loud
    [Arguments("=SUM(+IF(A1:A3>0,A1:A3))", "14")] // today #VALUE! — loud
    [Arguments("=SUM(-(+A1:A3))", "-14")] // today #VALUE! — loud
    public async Task AUnaryPlusOverAComputedArray_StreamsWhole_ArrayEnteredColumn(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11: these three are the rows where the columns SPLIT — plain entry answers
        // #VALUE! for all three (the classic implicit intersection of an operator result), CSE entry answers
        // 28 / 14 / -14. The CSE column is the target. SUM(-(+A1:A3)) is the row Phase 8 documented as a gap
        // while '+' was opaque (ElementwiseLiftingTests, "SUM(-(+A1:A3)) = -6" on its 1, 2, 3 grid); on this
        // grid the oracle's CSE column is -14, the same as SUM(-A1:A3) (#VALUE! plain / -14 CSE), so a
        // transparent '+' under a lifted '-' lifts exactly as if the '+' were not there.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    // ================================================================================================
    // A defined name whose definition is a computed array
    // ================================================================================================

    [Test]
    [Arguments("=SUM(ProdName)", "14")] // today 5 — silent
    [Arguments("=ROWS(ProdName)", "2")] // today 1 — silent
    [Arguments("=SUM(OpName)", "28")] // today #VALUE! — loud
    public async Task ADefinedNameOverAComputedArray_StreamsWholeToItsConsumer(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11: 14 / 2 / 28 in BOTH modes. Today ResolveNameShape answers Opaque for a
        // definition that is not a reference and EvaluateDefinition's CaptureValue collapses it — the same
        // site as a LET binding, one level up (NamedReferences.cs, EvaluateDefinition).
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    // ================================================================================================
    // The criteria gate sees through every binding site
    // ================================================================================================

    [Test]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,\">0\"))")] // today 1 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),SUMIF(f,\">0\"))")] // today 5 — silent
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),COUNTBLANK(f))")] // today 0 — silent
    [Arguments("=COUNTIF(ProdName,\">0\")")] // today 1 — silent
    [Arguments("=COUNTIF(CHOOSE(1,FILTER(A1:A3,A1:A3>0)),\">0\")")] // today 1 — silent
    [Arguments("=COUNTIF(+FILTER(A1:A3,A1:A3>0),\">0\")")] // today 1 — silent
    public async Task AnArrayBinding_InACriteriaSlot_IsRefused(string formula)
    {
        // Oracle 26.6.0, 2026-09-11: #REF! in BOTH modes for every row — a producer in a range slot is the
        // shape with no mode ambiguity (CriteriaComputedArgumentTests' header), and a binding does not change
        // that. Today every row answers a count or sum over the ONE collapsed element, silently.
        await Assert.That(Calc(formula)).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    [Arguments("=COUNTIF(LET(r,A1:A3,r*1),\">0\")")] // today 0 — silent
    [Arguments("=SUMIF(LET(r,A1:A3,r*1),\">0\")")] // today 0 — silent
    [Arguments("=LET(r,A1:A3*1,COUNTIF(r,\">0\"))")] // today 0 — silent
    [Arguments("=LET(r,A1:A3*1,SUMIF(r,\">0\"))")] // today 0 — silent
    public async Task ALetBoundOperatorArray_InACriteriaSlot_IsRefused_ArrayEnteredColumn(
        string formula
    )
    {
        // Oracle 26.6.0, 2026-09-11: #VALUE! plain / #REF! CSE for all four (the columns split, as they do
        // for every COMPUTED array in a range slot; the CSE column is the target). These are Phase 11a's
        // standing limit, pinned at 0 in CriteriaComputedArgumentTests until this phase flipped them.
        await Assert.That(Calc(formula)).IsEqualTo(ErrorValue.Reference);
    }

    // ================================================================================================
    // Guards — GREEN today and must stay green
    // ================================================================================================

    [Test]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),f)", "5")]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),f+1)", "6")]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),f)*2", "10")]
    [Arguments("=CHOOSE(1,FILTER(A1:A3,A1:A3>0))", "5")]
    [Arguments("=+FILTER(A1:A3,A1:A3>0)", "5")]
    [Arguments("=ProdName", "5")]
    public async Task ABareBindingInACell_ShowsItsTopLeft(string formula, string oracle)
    {
        // Oracle 26.6.0, 2026-09-11: 5 / 6 / 10 / 5 / 5 / 5 in BOTH modes (H20, outside A1:A3, so this is
        // the top-left rule and not an intersection). GUARD: the cell boundary (Workbook.cs, "CAPTURE, not
        // plain Evaluate") keeps showing the top-left, and the SCALAR path of every binding site must keep
        // answering the top-left when the whole formula is not array-eligible — f+1 is top-left + 1, not an
        // array. Read through the cell, not through Expression.Evaluate: the two paths have disagreed before
        // in this project (INDIRECT measured 14 through the cell and #REF! bare).
        await Assert.That(InCell(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    public async Task ALetBinding_IsEvaluatedOnce_EvenWhenItIsAnArray()
    {
        // Oracle 26.6.0, 2026-09-11: 0 in BOTH modes. GUARD: a LET binding is evaluated ONCE (the LET page's
        // own sentence); an array binding must be built once and read twice, never rebuilt with a second
        // RAND() draw. Green today for the wrong reason — the binding collapses to one scalar, and one
        // scalar read twice is trivially equal — and it must stay green when the binding carries the array.
        // Twenty fresh workbooks so a single lucky draw cannot hide a rebuild.
        for (var i = 0; i < 20; i++)
        {
            await Assert.That(Calc("=LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))")).IsEqualTo(0.0);
        }
    }

    [Test]
    [Arguments("=LET(n,ROWS(FILTER(A1:A3,A1:A3>0)),n)", "2")]
    [Arguments("=LET(f,A1:A3,SUM(FILTER(f,f>0)))", "14")]
    [Arguments("=SUM(+A1:A3)", "14")]
    public async Task AScalarOrRangeBinding_IsUnchanged(string formula, string oracle)
    {
        // Oracle 26.6.0, 2026-09-11: 2 / 14 / 14 in BOTH modes. GUARD: a scalar binding whose expression
        // merely CONTAINS a producer stays a scalar; a binding that IS a range keeps CaptureValue's reference
        // path (Phase 11a's Rule A); '+' over a bare reference still carries the REFERENCE — SUM(+A1:A3) is
        // 14 on this grid (6 on UnaryOperationTests' 1, 2, 3 grid and 356 on ElementwiseLiftingTests'
        // OnLengths grid), so making '+' transparent must not turn it into a lift.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }
}
