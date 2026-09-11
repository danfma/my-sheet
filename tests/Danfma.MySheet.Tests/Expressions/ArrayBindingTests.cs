using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Tests.Expressions.SelectionProducerFixture;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 11c (array bindings) — the acceptance pins for a computed array that crosses a BINDING site: a
/// <c>LET</c> binding, the chosen branch of <c>CHOOSE</c>, the operand of unary <c>+</c>, and a workbook defined
/// name whose definition is a computed-array expression. <b>Every acceptance row here arrived RED</b>, on the
/// commit that wrote the file for the LET/CHOOSE/<c>+</c>/name rows and on the commit that made each site carry
/// the array for the handful added alongside it: <c>NamedReferences.CaptureValue</c>'s fall-through evaluated a
/// non-reference binding as an ordinary scalar, so a producer collapsed to its top-left silently (5 where the
/// answer is 14, 1 where it is 2) and an operator over a range or an array <c>IF</c> collapsed to a loud
/// <c>#VALUE!</c>. A row's comment names the value it used to give beside the oracle's, and a row that is
/// green for the wrong reason says so in its own comment (the GUARD rows do).
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
/// columns agree except where the test's OWN comment names the split — every test whose name ends
/// <c>_ArrayEnteredColumn</c>, plus one row each in the two <c>AChainedRebindingOfAnIfOverARangeName…</c>
/// tests — and where they split <b>the CSE column is the target</b>, because the mini-CSE implements the
/// array-entered rule everywhere. Modes are never mixed inside one assertion. (A count of splitting rows is
/// deliberately not written here: it went stale twice while the file grew.)
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
    [Arguments("=LET(x,LEN(A1:A3),SUM(x))", "3")] // today 1 — silent (a LIFTED CALL binding, function-reference.md's LET row)
    [Arguments("=LET(x,ABS(A1:A3),SUM(x))", "14")] // today 5 — silent (same lifted-call shape, non-discriminating values)
    public async Task ALetBoundArray_StreamsWholeToItsConsumer(string formula, string oracle)
    {
        // Oracle 26.6.0, 2026-09-11: the value named, in BOTH entry modes, for every row. The two loud rows
        // are the binding shapes CaptureValue cannot even collapse — an operator over a range and an array IF
        // evaluate to #VALUE! on the scalar path — and the silent rows are a producer's FirstElement rule
        // applied one site too early: SUM([5]) = 5, ROWS([5]) = 1, SUM([5]*2) = 10. The last two rows pin
        // function-reference.md's "a lifted call" binding clause (GLM measured 2026-09-11, own probe copy):
        // LEN lifted over 5, 0, 9 is 1, 1, 1 (character counts of the digits), so SUM is 3, not 14 — the row
        // that actually discriminates a lift from a bare range passthrough; ABS is a lift too, but over
        // already-non-negative values so its own SUM cannot tell a lift from Rule A's range path, and is
        // pinned anyway because it is the literal example the doc's own comment names.
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
    [Arguments("=SUM(CHOOSE(1,A1:A3,FILTER(A1:A3,A1:A3>0)))", "14")]
    [Arguments("=SUM(CHOOSE(2,FILTER(A1:A3,A1:A3>0),A1:A3))", "14")]
    [Arguments("=ROWS(CHOOSE(1,A1:A3,FILTER(A1:A3,A1:A3>0)))", "3")]
    [Arguments("=ROWS(CHOOSE(2,FILTER(A1:A3,A1:A3>0),A1:A3))", "3")]
    public async Task AChosenBareRange_BesideAnArrayBranch_StillCarriesItsCells(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11 (own probe copy, H20, this file's fixture): 14 / 14 / 3 / 3 in BOTH modes.
        // These rows are where CHOOSE and a scalar-condition IF DIVERGE, and the divergence is in each
        // function's own scalar path, not in the mini-CSE: CHOOSE has captured a chosen range as a reference
        // VALUE since Onda 3, so the chosen A1:A3 streams the cells it denotes even though the sibling branch
        // is what made the node array-eligible, while SUM(IF(TRUE,A1:A3,SEQUENCE(3))) is #VALUE! because
        // RangeReference.Evaluate is (sweep item 32, unmoved — ArrayEvaluation's WrapScalar comment carries
        // that pair). Green before this phase too, through the collapsed reference value; the pin is here so
        // the CHOOSE arm cannot quietly adopt IF's answer for the shape. The index-2 ROWS row (final-review
        // fix wave) exists because SUM alone cannot discriminate: FILTER(...) sums to 14 and the sibling A1:A3
        // also sums to 14, so only ROWS(...) = 3, A1:A3's own row count, proves CHOOSE took the range branch.
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
        // Oracle 26.6.0, 2026-09-11: 14 / 2 / -14 in BOTH modes. Today (before this phase) UnaryOperation.Evaluate
        // captures the operand as a value, which for a producer is its FirstElement, and the '+' was opaque to
        // the probe; Phase 11c makes '+' transparent to the probe (ArrayEvaluation.cs), not lifted.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    [Arguments("=SUM(+(A1:A3*2))", "28")] // today #VALUE! — loud
    [Arguments("=SUM(+IF(A1:A3>0,A1:A3))", "14")] // today #VALUE! — loud
    [Arguments("=SUM(-(+A1:A3))", "-14")] // today #VALUE! — loud
    [Arguments("=SUM((+A1:A3)*B1:B3)", "32")] // today #VALUE! — loud
    public async Task AUnaryPlusOverAComputedArray_StreamsWhole_ArrayEnteredColumn(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11: these four are rows where the columns SPLIT — plain entry answers
        // #VALUE! for all of them (the classic implicit intersection of an operator result), CSE entry answers
        // 28 / 14 / -14 / 32. The CSE column is the target. SUM(-(+A1:A3)) is the row Phase 8 documented as a gap
        // while '+' was opaque (docs/workbook-and-expressions.md's "SUM(-(+A1:A3)) is -6" sentence, pinned by
        // UnaryOperationTests.Plus_OnRangeNode_IsStillARange's "=SUM(-(+A1:A3))" row on its own 1, 2, 3 grid);
        // on THIS file's fixture the oracle's CSE column is -14, the same as SUM(-A1:A3) (#VALUE! plain / -14
        // CSE), so a
        // transparent '+' under a lifted '-' lifts exactly as if the '+' were not there. The last row is the
        // same statement one operator over: a '+' over a bare RANGE inside an operator zips the cells
        // (5*1 + 0*2 + 9*3), which is only visible because the '+' no longer hides the range from the probe.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    [Arguments("=COUNTIF(+A1:A3,\">0\")", "2")]
    [Arguments("=SUMIF(+A1:A3,\">0\")", "14")]
    [Arguments("=AVERAGEIF(+A1:A3,\">0\")", "7")]
    [Arguments("=COUNTBLANK(+A1:A3)", "0")]
    [Arguments("=SUM(+A:A)", "28")]
    public async Task AUnaryPlusOverABareReference_IsStillAReferenceAtTheTopLevel(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11 (own probe copy, H20, this file's fixture): 2 / 14 / 7 / 0 in BOTH modes
        // for the first four, and ISREF(+A1:A3) is TRUE in both (pinned by the sibling test below) — so a '+'
        // over a bare reference DENOTES that reference where a consumer asks for one, and the criteria family
        // reads its cells exactly as it reads A1:A3's. GUARD, and the one that decides a design question the
        // phase plan did not: making the '+' transparent to the probe (item 11) would have turned all FIVE of
        // these shapes — COUNTIF/SUMIF/AVERAGEIF/COUNTBLANK to #REF! through CriteriaScan.RejectComputedArray
        // and ISREF to FALSE — five new divergences — had IsBareReferenceNode not been given its Plus arm,
        // which answers for the OPERAND. The contrast is one row over: COUNTIF(+FILTER(…),">0") is #REF!
        // (AnArrayBinding_InACriteriaSlot_IsRefused), because a producer under the '+' denotes no reference.
        //
        // The fifth row is the cost guard rather than the criteria gate: an OPEN range under '+' is REFUSED
        // by the probe, so the '+' stays the opaque scalar that carries the reference and SUM reads the
        // populated column — 5 + 0 + 9 + 7 + 7 = 28 on this fixture (A5 and A8 are 7, A7 is text and skipped),
        // which is what the oracle answers in both modes, and what SUM(A:A) answers without the '+'.
        await Assert.That(Calc(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    public async Task AUnaryPlusOverABareReference_IsStillAReferenceForIsref()
    {
        // Oracle 26.6.0, 2026-09-11: ISREF(+A1:A3) is TRUE in both modes — the fifth of the five shapes named
        // in the sibling test's comment above, pinned separately because a boolean result does not fit that
        // method's shared numeric/error Oracle() helper.
        await Assert.That(Calc("=ISREF(+A1:A3)") as bool?).IsTrue();
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
    // Guards — protect a shape that must not move; not every row below was green on ARRIVAL (four
    // weren't — each says so in its own comment), but every one must stay green from here on
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
        // own sentence); an array binding is built once at binding time (ArrayBindings.Capture, Let.TryBind)
        // and read twice through the scope, never rebuilt with a second RAND() draw. History: this row was
        // green BEFORE the binding carried the array, for the wrong reason — the binding collapsed to one
        // scalar, and one scalar read twice is trivially equal — which is why the counting row below exists:
        // it is the one that turns red under a rebuild. Twenty fresh workbooks so a single lucky draw
        // cannot hide one either.
        for (var i = 0; i < 20; i++)
        {
            await Assert.That(Calc("=LET(x,SEQUENCE(3,1,RAND(),0),SUM(x)-SUM(x))")).IsEqualTo(0.0);
        }
    }

    [Test]
    public async Task ALetArrayBinding_IsBuiltOnce_WhenTheNameIsReadTwice()
    {
        // Engine pin (no oracle column — a custom function has no oracle): the binding EXPRESSION of an array
        // binding runs exactly once however many times the name is read. TICK() counts its own calls, and
        // SEQUENCE(3,1,TICK(),1) is array-eligible, so this row enters the ARRAY branch of the binding
        // (ROWS(x) is 3 only when the operand, not its top-left, is bound: 1 + 2 + 3 = 6 from a first tick
        // of 1, so 3*10 + 6 = 36 pins both the shape and the value); the scalar-binding control below
        // enters the CaptureValue branch, which was already evaluate-once. Under the collapse (the mutation
        // in Phase 2's verification plan — item 5's array branch reverted to CaptureValue) this row answers
        // 11 (ROWS 1, SUM 1) with the count still 1, because the collapse evaluates once as well: the shape
        // is what proves the branch, the count what proves evaluate-once.
        var (workbook, sheet) = Named();
        var ticks = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++ticks);

        var array = ExpressionParser
            .Parse("=LET(x,SEQUENCE(3,1,TICK(),1),ROWS(x)*10+SUM(x))", sheet)
            .Evaluate(workbook)
            .AsObject();

        await Assert.That(array).IsEqualTo(36.0);
        await Assert.That(ticks).IsEqualTo(1);

        ticks = 0;
        var scalar = ExpressionParser
            .Parse("=LET(n,TICK(),n+n)", sheet)
            .Evaluate(workbook)
            .AsObject();

        await Assert.That(scalar).IsEqualTo(2.0);
        await Assert.That(ticks).IsEqualTo(1);
    }

    [Test]
    [Arguments("=SUM(CHOOSE(IF(RAND()<0.5,1,2),FILTER(A1:A3,A1:A3>0),SEQUENCE(3)))", 14.0, 6.0)]
    [Arguments("=SUM(-CHOOSE(IF(RAND()<0.5,1,2),FILTER(A1:A3,A1:A3>0),SEQUENCE(3)))", -14.0, -6.0)]
    public async Task AVolatileChooseIndex_NeverCollapsesTheBranchItTakes(
        string formula,
        double whenFirst,
        double whenSecond
    )
    {
        // 200 fresh workbooks, each drawing its own RAND(): the index picks branch 1 (FILTER, sum 14) or
        // branch 2 (SEQUENCE(3), sum 6), and NOTHING else may appear. 5 is FILTER's top-left and 1 is
        // SEQUENCE's — the two silent collapses this phase removes — so either of them here would mean the
        // chosen branch was read as a scalar, and a value belonging to the OTHER branch than the one the
        // index took would mean the index was drawn twice (a consumer falling back into Choose.Evaluate
        // after the probe had promised an array). Both buckets must actually occur, or the run proves
        // nothing about the branch it never entered: that is the exact shape of Task 8's volatility defect
        // one binding site over (lessons.md, 2026-09-11 — "a volatility test must exercise the shape that
        // reaches the code under test"), where a mixed IF found a producer collapsing in 100 of 400 seeds.
        // The second row wraps the same node in a lifted '-' so the taken branch is read through an operand
        // rather than by the consumer directly. Measured buckets on this tree: 94/106 and 103/97.
        var buckets = new Dictionary<object, int>();

        for (var i = 0; i < 200; i++)
        {
            var value = Calc(formula) ?? "null";
            buckets[value] = buckets.GetValueOrDefault(value) + 1;
        }

        await Assert.That(buckets.Count).IsEqualTo(2);
        await Assert.That(buckets.GetValueOrDefault(whenFirst)).IsGreaterThan(0);
        await Assert.That(buckets.GetValueOrDefault(whenSecond)).IsGreaterThan(0);
    }

    [Test]
    [Arguments("=LET(x,A1:A3*2,x)", "10")] // today #VALUE! — loud
    [Arguments("=LET(x,IF(A1:A3>0,A1:A3),x)", "5")] // today #VALUE! — loud
    [Arguments("=OpName", "10")] // today #VALUE! — loud
    [Arguments("=+(A1:A3*2)", "10")] // today #VALUE! — loud
    [Arguments("=CHOOSE(1,A1:A3*2)", "10")] // today #VALUE! — loud
    public async Task ABareOperatorBindingInACell_ShowsItsTopLeft_ArrayEnteredColumn(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11: CSE 10 / 5 / 10 / 10 / 10; PLAIN intersects per row for the LET and
        // operator shapes (10 in H1, #VALUE! in H20 — the controller measured the first two in H1 and H5 and
        // this task re-measured all five in H20), so the columns split and the CSE column is the target. The
        // exception is =OpName, which is 10 in BOTH columns.
        //
        // The SCALAR reading of an array-eligible binding is the built operand's TOP-LEFT
        // (ArrayBindings.Binding.TopLeft through ArrayEvaluation.FirstElement), the same rule a bare producer
        // follows — not Evaluate's #VALUE!, which is what an operator over a range answers on the scalar path
        // and what MySheet answered for all five before this phase. One rule at all four binding sites: the
        // last three rows are the defined name (NamedReferences.EvaluateDefinition), the unary '+'
        // (UnaryOperation.Evaluate) and CHOOSE's chosen branch (Choose.Evaluate), each routed through the same
        // ArrayBindings.Capture that Let uses. Read through the cell: these are the rows that prove the cell
        // boundary needs no change for an operator binding at any of the four.
        await Assert.That(InCell(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    public async Task ABareOperatorArrayInACell_StaysTheBoundaryDecision()
    {
        // CONTROL, an engine pin rather than an oracle row: bare =A1:A3*2 in a cell is still #VALUE!
        // (CellBoundaryIntersectionTests — the cell-boundary array gap is a separate sweep decision). The
        // top-left rule above is a rule for a BINDING's scalar reading, applied by NameReference.Evaluate to
        // a bound name; Workbook's boundary keeps calling plain CaptureValue and never applies it to a bare
        // operator node, so a LET around the same expression is 10 (the row above) while the bare node is
        // not.
        await Assert.That(InCell("=A1:A3*2")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),f)", "5")]
    [Arguments("=LET(f,FILTER(A1:A3,A1:A3>0),f+1)", "6")]
    public async Task TheCellBoundary_NeverSeesAnArray_BecauseALetNodeEvaluatesToItsTopLeft(
        string formula,
        string oracle
    )
    {
        // WHY Workbook's cell boundary needs no change for an array binding (Phase 11c item 8): the bare
        // formula reaches the boundary as a Let NODE, and NamedReferences.CaptureValue on a Let node is
        // Let.Evaluate, whose body reads the bound name through NameReference.Evaluate — the operand's
        // top-left, a plain scalar. So the value the boundary receives is neither a reference (nothing to
        // intersect) nor an array (nothing to spill): the three assertions pin that the capture is not a
        // reference value, that the expression path and the cell path agree, and that both are the
        // oracle's top-left (5 / 6, Aspose.Cells 26.6.0, 2026-09-11, both modes).
        var (workbook, sheet) = Named();
        var expression = ExpressionParser.Parse(formula, sheet);
        var captured = NamedReferences.CaptureValue(
            expression,
            new EvaluationContext(workbook, "Sheet1", "H20")
        );

        await Assert.That(captured.TryGetReference(out _)).IsFalse();
        await Assert.That(captured.AsObject()).IsEqualTo(Oracle(oracle));
        await Assert.That(InCell(formula)).IsEqualTo(Oracle(oracle));
    }

    // A range NAME for the seam rows below: IF(TRUE,Rng,0) passes the name's reference VALUE out of
    // If.Evaluate without answering TryResolveReference itself, which is the one shape where
    // ArrayBindings.Shape (the probe's stand-in) and ArrayBindings.Capture (the build) classify a binding
    // differently — a scalar to the probe, a range-bound name to the build.
    private static object? WithRng(string formula)
    {
        var (workbook, sheet) = Named();
        workbook.DefineName("Rng", "Sheet1!$A$1:$A$3");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    [Arguments("=LET(a,IF(TRUE,Rng,0),b,a,SUM(b*1))", "14")]
    [Arguments("=SUM(LET(a,IF(TRUE,Rng,0),b,a,b*1)*B1:B3)", "32")]
    public async Task AChainedRebindingOfAnIfOverARangeName_ResolvesOnTheBuildSide(
        string formula,
        string oracle
    )
    {
        // Oracle 26.6.0, 2026-09-11 (own probe copy, this fixture): 14 in BOTH modes; #VALUE! plain / 32
        // CSE (the columns split, as they do for every operator over a range — the CSE column is the
        // target). The build side is where the seam resolves right: a is Capture'd as the reference value
        // If.Evaluate carries, b rebinds it, and b*1 zips over the three cells — for the first row
        // directly (SUM's argument is b*1, a Binary the probe answers for through the name's REAL scope),
        // and for the second because B1:B3 makes the outer Binary an array whatever the probe said about
        // the Let, so the build runs. The second row MOVED with Phase 11c's Let arm: #VALUE! before (the
        // Let node was an opaque scalar to the build too), 32 now.
        await Assert.That(WithRng(formula)).IsEqualTo(Oracle(oracle));
    }

    [Test]
    [Arguments("=SUM(LET(a,IF(TRUE,Rng,0),b,a,b*1))", "#VALUE!")]
    [Arguments("=COUNTIF(LET(a,IF(TRUE,Rng,0),b,a,b*1),\">0\")", "0")]
    public async Task AChainedRebindingOfAnIfOverARangeName_IsOpaqueToTheProbe_KnownDivergence(
        string formula,
        string oracle
    )
    {
        // KNOWN DIVERGENCE, unchanged by Phase 11c (measured identical on 6ad7cea, before the Let arm):
        // the oracle (26.6.0, 2026-09-11, own probe copy) answers 14 in BOTH modes for the first row and
        // #VALUE! plain / #REF! CSE for the second; the engine answers #VALUE! and 0. The probe's stand-in
        // for a binding resolves reference-ness through TryResolveReference, which IF does not answer, so
        // a is a scalar in the probe's scope, the Let is "not an array" to the consumer's gate, and the
        // consumer keeps its scalar path (SUM evaluates the Let: a reference value times 1 is #VALUE!) or
        // its cursor (COUNTIF opens the one collapsed element: 0). The probe is only ever MORE conservative
        // than the build here, never the reverse — the reason the invariant "eligible iff the build is an
        // array" is not broken in the direction a consumer could observe as a double evaluation. Owned by
        // the "IF returns a reference" question (sweep item 32), not by this phase; a pin so the seam
        // ArrayBindings.Shape documents is not a silent one.
        await Assert.That(WithRng(formula)).IsEqualTo(Oracle(oracle));
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
