using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 11a Rule A (item 8 of the compatibility sweep): <b>a defined name in an array position is whatever
/// it is bound to; at a consumer's top level a bare name is a reference and takes the reference path, exactly
/// as a bare literal range does.</b>
/// <para>
/// Oracle: <b>Aspose.Cells 26.6.0</b>, measured 2026-09-10 on this exact fixture. Every golden number below
/// is the <b>ARRAY-ENTERED</b> column (<c>Cell.SetArrayFormula(f, 1, 1)</c>), which is the mode the mini-CSE
/// implements; where the PLAIN (typed) column differs it is named in the test's comment and never compared
/// against the array-entered one. For the lifted shapes plain is <c>#VALUE!</c> (implicit intersection) and
/// carries no signal.
/// </para>
/// <para>
/// Two halves. The <b>acceptance</b> pins are RED at <c>0b93d66</c> — <c>ArrayEvaluation.Probe</c> has no
/// <c>NameReference</c> arm, so a name in a nested array position is an opaque scalar and
/// <c>TryBuildOperand</c> broadcasts the reference VALUE (which compares as TRUE, hence the plausible 1s).
/// Each states the invariant, not just a number: the name must answer what its LITERAL twin answers, so the
/// pin survives a later change to the literal's own semantics. The <b>must-not-move</b> pins are GREEN today
/// and must stay green: <c>NameReference</c> derives from <c>Expression</c>, not <c>Reference</c>, so once
/// <c>IsArrayEligible</c> says TRUE for a range-bound name, every consumer gate written
/// <c>is not Reference &amp;&amp; IsArrayEligible(…)</c> would route a BARE top-level name into the mini-CSE.
/// Measured on the designer's no-guard prototype, that regresses fifteen shapes; the pins below are those
/// shapes, and they are the only thing standing between a later "simplification" of a gate and a silent loss.
/// </para>
/// </summary>
public class DefinedNameArrayEligibilityTests
{
    // Rule A's fixture, verbatim. A1:A3 = 5, 0, 9; B1:B3 = 1, 2, 3; D1:D3 = 1, 22, 333; F1:F2 = 1, 2 and
    // F3 = SUBTOTAL(9,F1:F2) = 3 (the nested-subtotal probe). A2 = 0 is what makes the difference between
    // "one reference value, truthy" and "three elements" visible as a NUMBER rather than an error.
    // Names are workbook-level, so every one is sheet-qualified (Workbook.DefineName rejects the bare form);
    // GhostName points at a sheet that does not exist, MyCol is the open-range (cost-guard) twin of Rng,
    // MyCell the single-cell one, UnN the union one and Threshold a constant.
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);
        sheet["B1"] = new NumberValue(1);
        sheet["B2"] = new NumberValue(2);
        sheet["B3"] = new NumberValue(3);

        sheet["D1"] = new NumberValue(1);
        sheet["D2"] = new NumberValue(22);
        sheet["D3"] = new NumberValue(333);

        sheet["F1"] = new NumberValue(1);
        sheet["F2"] = new NumberValue(2);
        sheet["F3"] = ExpressionParser.Parse("=SUBTOTAL(9,F1:F2)", sheet);

        workbook.DefineName("Rng", "Sheet1!$A$1:$A$3");
        workbook.DefineName("Wide", "Sheet1!$A$1:$B$3");
        workbook.DefineName("MyName", "Sheet1!$D$1:$D$3");
        workbook.DefineName("Nested", "Sheet1!$F$1:$F$3");
        workbook.DefineName("MyCell", "Sheet1!$A$2");
        workbook.DefineName("MyCol", "Sheet1!$A:$A");
        workbook.DefineName("GhostName", "Ghost!$A$1:$A$3");
        workbook.DefineName("UnN", "(Sheet1!$A$1:$A$2,Sheet1!$A$3:$A$3)");
        workbook.DefineName("Threshold", new NumberValue(4));

        return (workbook, sheet);
    }

    // The 2-D error variant of Grid(): A3 = #DIV/0! and B1 = #N/A, so Wide (A1:B3) holds TWO errors in
    // different columns and the order in which the engine scans it decides which one wins.
    private static (Workbook Workbook, Sheet Sheet) WithErrors()
    {
        var (workbook, sheet) = Grid();
        sheet["A3"] = ExpressionParser.Parse("=1/0", sheet);
        sheet["B1"] = ExpressionParser.Parse("=NA()", sheet);
        return (workbook, sheet);
    }

    private static object? OnGrid(string formula)
    {
        var (workbook, sheet) = Grid();
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static object? OnErrors(string formula)
    {
        var (workbook, sheet) = WithErrors();
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // ================================================================================================
    // ACCEPTANCE pins (RED at 0b93d66) — the name must answer its literal twin's answer.
    // In every one of them the literal twin ALREADY answers the golden value here (measured 2026-09-10),
    // so the name-equals-literal assertion is itself the red one.
    // ================================================================================================

    [Test]
    public async Task Count_OfANameComparedToBlank_CountsEveryElement()
    {
        // Aspose 26.6.0 array-entered: 3 (plain: 0). MySheet 0b93d66: name 1, literal 3 — the SILENT class:
        // the name broadcasts as ONE truthy reference value, so COUNT sees a single element.
        var name = OnGrid("=COUNT((Rng<>\"\")*1)");
        await Assert.That(name).IsEqualTo(OnGrid("=COUNT((A1:A3<>\"\")*1)"));
        await Assert.That(Num(name)).IsEqualTo(3.0);
    }

    [Test]
    public async Task Sum_OfANameComparedToZero_SumsTheElementwiseFlags()
    {
        // Aspose 26.6.0 array-entered: 2 (plain: #VALUE!). MySheet 0b93d66: name 1, literal 2 — SILENT: the
        // "1" is the truthy reference value, not the two non-zero cells, and no error announces it.
        var name = OnGrid("=SUM((Rng<>0)*1)");
        await Assert.That(name).IsEqualTo(OnGrid("=SUM((A1:A3<>0)*1)"));
        await Assert.That(Num(name)).IsEqualTo(2.0);
    }

    [Test]
    public async Task Count_OfANameTimesOne_CountsThreeElements()
    {
        // Aspose 26.6.0 array-entered: 3 (plain: 0). MySheet 0b93d66: name 0 — a NUMBER, not an error: the
        // reference value times one is not numeric, so COUNT counts nothing. The most silent row of Rule A.
        var name = OnGrid("=COUNT(Rng*1)");
        await Assert.That(name).IsEqualTo(OnGrid("=COUNT(A1:A3*1)"));
        await Assert.That(Num(name)).IsEqualTo(3.0);
    }

    [Test]
    public async Task SmallAndMin_OfAnIfFilteredName_SeeTheFilteredElements()
    {
        // Aspose 26.6.0 array-entered: 5 and 5 (plain: 5 and #VALUE!). MySheet 0b93d66: 0 and 0 for the
        // name (the opaque 1x1 collapses the filter), 5 and 5 for the literal twins. SILENT both times —
        // and 0 is a legal answer over this fixture, which is what makes the pin necessary.
        var small = OnGrid("=SMALL(IF(Rng>0,Rng),1)");
        await Assert.That(small).IsEqualTo(OnGrid("=SMALL(IF(A1:A3>0,A1:A3),1)"));
        await Assert.That(Num(small)).IsEqualTo(5.0);

        var min = OnGrid("=MIN(IF(Rng>0,Rng))");
        await Assert.That(min).IsEqualTo(OnGrid("=MIN(IF(A1:A3>0,A1:A3))"));
        await Assert.That(Num(min)).IsEqualTo(5.0);
    }

    [Test]
    public async Task Index_OfANameTimesTwo_IndexesTheComputedVector()
    {
        // Aspose 26.6.0 array-entered: 18 (plain: 18 — both modes agree). MySheet 0b93d66: name #REF!,
        // literal 18. This is the exact operand shape Phase 7's FILTER hands to TryBuildOperand, which is
        // why it is in the unblocking slice at all.
        var name = OnGrid("=INDEX(Rng*2,3)");
        await Assert.That(name).IsEqualTo(OnGrid("=INDEX(A1:A3*2,3)"));
        await Assert.That(Num(name)).IsEqualTo(18.0);
    }

    [Test]
    public async Task Sum_OfIfOverANameWithRowOfTheName_AddsTheMatchingRowNumbers()
    {
        // Aspose 26.6.0 array-entered: 4 (plain: #VALUE!). MySheet 0b93d66: name 1, literal 4 — SILENT.
        // ROW(Rng) already resolves the name through ResolvePositionRange (Phase 1); the condition does not.
        var name = OnGrid("=SUM(IF(Rng>0,ROW(Rng)))");
        await Assert.That(name).IsEqualTo(OnGrid("=SUM(IF(A1:A3>0,ROW(A1:A3)))"));
        await Assert.That(Num(name)).IsEqualTo(4.0);
    }

    [Test]
    public async Task Sum_MixingANameWithItsOwnPositionVector_MultipliesElementwise()
    {
        // Aspose 26.6.0 array-entered: 32 and 2 (plain: #VALUE! for both). MySheet 0b93d66: the first is
        // #VALUE! (a position VECTOR times an opaque scalar), the second a SILENT 1; literals 32 and 2.
        var product = OnGrid("=SUM(ROW(Rng)*Rng)");
        await Assert.That(product).IsEqualTo(OnGrid("=SUM(ROW(A1:A3)*A1:A3)"));
        await Assert.That(Num(product)).IsEqualTo(32.0);

        var twoFlags = OnGrid("=SUM((Rng<>\"\")*(Rng<>0))");
        await Assert.That(twoFlags).IsEqualTo(OnGrid("=SUM((A1:A3<>\"\")*(A1:A3<>0))"));
        await Assert.That(Num(twoFlags)).IsEqualTo(2.0);
    }

    [Test]
    public async Task LiftedShapes_OverAName_AnswerWhatTheLiteralAnswers()
    {
        // Aspose 26.6.0 array-entered: 20, -14, 28, 3 (plain: #VALUE! on all four — implicit intersection,
        // no signal). MySheet 0b93d66: #VALUE! on all four for the name; the literal twins already answer
        // 20, -14, 28, 3. ElementwiseLiftingTests owns the MyName (1, 22, 333) copy of this row, which
        // item 4 of this phase flips.
        var sum = OnGrid("=SUM(Rng+B1:B3)");
        await Assert.That(sum).IsEqualTo(OnGrid("=SUM(A1:A3+B1:B3)"));
        await Assert.That(Num(sum)).IsEqualTo(20.0);

        var negated = OnGrid("=SUM(-Rng)");
        await Assert.That(negated).IsEqualTo(OnGrid("=SUM(-A1:A3)"));
        await Assert.That(Num(negated)).IsEqualTo(-14.0);

        var doubled = OnGrid("=SUM(Rng*2)");
        await Assert.That(doubled).IsEqualTo(OnGrid("=SUM(A1:A3*2)"));
        await Assert.That(Num(doubled)).IsEqualTo(28.0);

        var lengths = OnGrid("=SUM(LEN(Rng))");
        await Assert.That(lengths).IsEqualTo(OnGrid("=SUM(LEN(A1:A3))"));
        await Assert.That(Num(lengths)).IsEqualTo(3.0);
    }

    [Test]
    public async Task SumproductAndIndex_OverANameBuiltArray_SeeThreeElements()
    {
        // Aspose 26.6.0: 2 and 0 in BOTH modes. MySheet 0b93d66: SUMPRODUCT a SILENT 1, INDEX #REF!;
        // literals 2 and 0. SUMPRODUCT's own opt-in factory (OpenArrayOrRange) is untouched by this rule —
        // what changes is only that the OPERAND under it becomes three elements instead of one.
        var product = OnGrid("=SUMPRODUCT((Rng<>0)*1)");
        await Assert.That(product).IsEqualTo(OnGrid("=SUMPRODUCT((A1:A3<>0)*1)"));
        await Assert.That(Num(product)).IsEqualTo(2.0);

        var second = OnGrid("=INDEX((Rng<>0)*1,2)");
        await Assert.That(second).IsEqualTo(OnGrid("=INDEX((A1:A3<>0)*1,2)"));
        await Assert.That(Num(second)).IsEqualTo(0.0);
    }

    [Test]
    public async Task Aggregate_OverANameBuiltArray_SelectsOverTheElements()
    {
        // Aspose 26.6.0: 0 and 9 in BOTH modes. MySheet 0b93d66: AGGREGATE(15,6,…) a SILENT 1 (the truthy
        // reference value is the only element, so the 1st smallest is it), AGGREGATE(14,6,…) #VALUE!;
        // literals 0 and 9.
        var smallest = OnGrid("=AGGREGATE(15,6,(Rng<>0)*1,1)");
        await Assert.That(smallest).IsEqualTo(OnGrid("=AGGREGATE(15,6,(A1:A3<>0)*1,1)"));
        await Assert.That(Num(smallest)).IsEqualTo(0.0);

        var largest = OnGrid("=AGGREGATE(14,6,Rng*1,1)");
        await Assert.That(largest).IsEqualTo(OnGrid("=AGGREGATE(14,6,A1:A3*1,1)"));
        await Assert.That(Num(largest)).IsEqualTo(9.0);
    }

    [Test]
    public async Task Sum_ComparingARangeNameToASingleCellName_BroadcastsTheCellsValue()
    {
        // Aspose 26.6.0 array-entered: 2 (plain: #VALUE!). MySheet 0b93d66: name 1, literal 2 — SILENT.
        // MyCell = $A$2 = 0, so the correct answer counts A1 and A3. The single-cell name must broadcast
        // its VALUE (Scalar), not its reference; the RANGE name is what has to become three elements.
        var name = OnGrid("=SUM((Rng>MyCell)*1)");
        await Assert.That(name).IsEqualTo(OnGrid("=SUM((A1:A3>$A$2)*1)"));
        await Assert.That(Num(name)).IsEqualTo(2.0);
    }

    [Test]
    public async Task TwoDimensionalName_InAnArrayPosition_StreamsTheWholeRectangle()
    {
        // Aspose 26.6.0 array-entered: 2 and 46 (plain: #VALUE! for both). MySheet 0b93d66: a SILENT 1 and
        // #VALUE!; literals 2 and 46. Wide = A1:B3, so a 2-D name must resolve exactly like the 2-D literal
        // (5 and 9 are > 4; 5*1 + 0*2 + 9*3 + 1*1 + 2*2 + 3*3 = 46 with B1:B3 broadcast across both columns).
        var flags = OnGrid("=SUM((Wide>4)*1)");
        await Assert.That(flags).IsEqualTo(OnGrid("=SUM((A1:B3>4)*1)"));
        await Assert.That(Num(flags)).IsEqualTo(2.0);

        var product = OnGrid("=SUM(Wide*B1:B3)");
        await Assert.That(product).IsEqualTo(OnGrid("=SUM(A1:B3*B1:B3)"));
        await Assert.That(Num(product)).IsEqualTo(46.0);
    }

    [Test]
    public async Task Subtotal_OverANameBuiltArray_RejectsItLikeTheLiteral()
    {
        // The ref-slot REJECTION rides along on eligibility, and it is a BEHAVIOUR CHANGE, not a fix of a
        // silent number: SUBTOTAL's ref slot takes references only, so once (Rng<>0)*1 is array-eligible
        // AggregateCodes.Feed refuses it exactly as it refuses (A1:A3<>0)*1 today.
        // Aspose 26.6.0: #VALUE! in BOTH modes. MySheet 0b93d66: name a SILENT 1, literal #VALUE!.
        var name = OnGrid("=SUBTOTAL(9,(Rng<>0)*1)");
        await Assert.That(name).IsEqualTo(OnGrid("=SUBTOTAL(9,(A1:A3<>0)*1)"));
        await Assert.That(name).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task GhostSheetName_InAnArrayPosition_StreamsThePerElementRefError()
    {
        // Correction B2 of the sweep: a name whose SHEET is missing still resolves to a RangeReference, so
        // the value vector streams the per-element #REF! — SUM propagates it and COUNT counts no numbers.
        // Aspose 26.6.0: #REF! in both modes for the SUM, 0 in both modes for the COUNT. MySheet 0b93d66:
        // a SILENT 1 for both (the unresolvable name broadcast as one truthy value); literals #REF! and 0.
        var sum = OnGrid("=SUM((GhostName<>0)*1)");
        await Assert.That(sum).IsEqualTo(OnGrid("=SUM((Ghost!A1:A3<>0)*1)"));
        await Assert.That(sum).IsEqualTo(ErrorValue.Reference);

        var count = OnGrid("=COUNT((GhostName<>\"\")*1)");
        await Assert.That(count).IsEqualTo(OnGrid("=COUNT((Ghost!A1:A3<>\"\")*1)"));
        await Assert.That(Num(count)).IsEqualTo(0.0);
    }

    [Test]
    public async Task NameInAnArrayPosition_OnTheErrorFixture_ReportsTheLiteralsError()
    {
        // Aspose 26.6.0 array-entered: #N/A, #DIV/0!, 10 (plain: #VALUE!, #VALUE!, 0 — a different mode,
        // never compared against these). MySheet 0b93d66: #VALUE!, #VALUE!, #DIV/0!; the literal twins
        // already answer #N/A, #DIV/0! and 10.
        // Wide = A1:B3 holds #N/A at B1 and #DIV/0! at A3; the elementwise stream is row-major, so #N/A
        // wins for Wide while the Rng column sees only #DIV/0!. IFERROR then flattens both to 0 → 10.
        var wide = OnErrors("=SUM(Wide*1)");
        await Assert.That(wide).IsEqualTo(OnErrors("=SUM(A1:B3*1)"));
        await Assert.That(wide).IsEqualTo(ErrorValue.NotAvailable);

        var column = OnErrors("=SUM(Rng*1)");
        await Assert.That(column).IsEqualTo(OnErrors("=SUM(A1:A3*1)"));
        await Assert.That(column).IsEqualTo(ErrorValue.DivByZero);

        var guarded = OnErrors("=SUM(IFERROR(Wide,0))");
        await Assert.That(guarded).IsEqualTo(OnErrors("=SUM(IFERROR(A1:B3,0))"));
        await Assert.That(Num(guarded)).IsEqualTo(10.0);
    }

    // ================================================================================================
    // MUST-NOT-MOVE pins — GREEN today and after. Every one of them REGRESSED on the designer's
    // prototype when the three top-level guards (TryStream, Index.TryResolveReference,
    // NumericAggregation's default arm) were left out, because NameReference is not a Reference and
    // IsArrayEligible now says TRUE for it. They are LOAD-BEARING: they are the only evidence that the
    // guards exist and work, and each one's failure mode is a plausible answer, not a crash.
    // ================================================================================================

    [Test]
    public async Task Subtotal_OverABareName_StaysOnTheReferencePath()
    {
        // LOAD-BEARING. Without the TryStream guard the bare name enters the mini-CSE and AggregateCodes.Feed
        // — whose ref slot takes references only — REJECTS it: measured on the no-guard prototype,
        // SUBTOTAL(9,Rng) 14 → #VALUE!, SUBTOTAL(3,Rng) 3 → #VALUE!, AGGREGATE(9,4,Rng) 14 → #VALUE!.
        // Aspose 26.6.0: 14, 3, 14 in BOTH modes.
        await Assert.That(Num(OnGrid("=SUBTOTAL(9,Rng)"))).IsEqualTo(14.0);
        await Assert.That(Num(OnGrid("=SUBTOTAL(3,Rng)"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=AGGREGATE(9,4,Rng)"))).IsEqualTo(14.0);
    }

    [Test]
    public async Task NestedSubtotalSkip_ThroughABareName_Survives()
    {
        // LOAD-BEARING, and the worst failure mode in the set: Nested = F1:F3 with F3 = SUBTOTAL(9,F1:F2),
        // and option 0 must SKIP that nested subtotal — which only the reference path can see, because the
        // skip is a property of the CELL, not of its value. On the no-guard prototype AGGREGATE(14,0,…,1)
        // went 2 → 3, SILENTLY counting F3: the nested skip was simply lost. AGGREGATE(9,0,Nested) and
        // SUBTOTAL(9,Nested) went 3 → #VALUE! there.
        // Aspose 26.6.0: 3, 3, 2 in BOTH modes.
        await Assert.That(Num(OnGrid("=AGGREGATE(9,0,Nested)"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=SUBTOTAL(9,Nested)"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=AGGREGATE(14,0,Nested,1)"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task OrderSelection_OverABareName_ReadsTheCells()
    {
        // A2 = 0, so the smallest of Rng is 0 — a value a broken path can also produce, which is why these
        // sit next to the SUBTOTAL pins rather than alone. Aspose 26.6.0: 0 and 0 in BOTH modes.
        await Assert.That(Num(OnGrid("=AGGREGATE(15,6,Rng,1)"))).IsEqualTo(0.0);
        await Assert.That(Num(OnGrid("=SMALL(Rng,1)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Index_OverABareName_StillReturnsAReference()
    {
        // LOAD-BEARING. Index.TryResolveReference refuses an array-eligible argument, so without its guard
        // INDEX over a bare name stops being a REFERENCE and every consumer of that reference breaks.
        // Measured on the no-guard prototype: SUM(A1:INDEX(Rng,3)) 14 → #REF!, ROW(INDEX(Rng,2)) 2 →
        // #VALUE!, ISREF(INDEX(Rng,2)) TRUE → FALSE, OFFSET(INDEX(Rng,1),1,0) 0 → #REF!.
        // Aspose 26.6.0: 9, 14, 2, TRUE, 0 in BOTH modes.
        await Assert.That(Num(OnGrid("=INDEX(Rng,3)"))).IsEqualTo(9.0);
        await Assert.That(Num(OnGrid("=SUM(A1:INDEX(Rng,3))"))).IsEqualTo(14.0);
        await Assert.That(Num(OnGrid("=ROW(INDEX(Rng,2))"))).IsEqualTo(2.0);
        await Assert.That(OnGrid("=ISREF(INDEX(Rng,2))") as bool?).IsTrue();
        await Assert.That(Num(OnGrid("=OFFSET(INDEX(Rng,1),1,0)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task RangeConsumers_OverABareName_AreUnchanged()
    {
        // The rest of the top-level corpus, all probed unchanged on the prototype WITH the guards. ROWS is
        // shape-only (3 either way) and is here to pin that the shape answer does not shift either;
        // COUNT(ROW(Rng)) = 3 is the element count of the position vector, which MiniCseConsumerTests owns
        // and this phase must not move (Aspose plain answers 1 there — implicit intersection; the
        // array-entered column is 3, and 3 is what MySheet answers).
        // Aspose 26.6.0 array-entered: 32, 2, 4, 3, 3, 14.
        await Assert.That(Num(OnGrid("=SUMPRODUCT(Rng,B1:B3)"))).IsEqualTo(32.0);
        await Assert.That(Num(OnGrid("=COUNTIF(Rng,\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUMIFS(B1:B3,Rng,\">0\")"))).IsEqualTo(4.0);
        await Assert.That(Num(OnGrid("=ROWS(Rng)"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=COUNT(ROW(Rng))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=SUM(IF(TRUE,Rng,B1:B3))"))).IsEqualTo(14.0);
    }

    [Test]
    public async Task ScalarAndUnknownNames_InAnArrayPosition_AreUnchanged()
    {
        // The three non-range name shapes, which the new arm must classify as Scalar (a constant, a single
        // cell) or Opaque (an unknown name, whose #NAME? must still flow through the scalar path).
        // Aspose 26.6.0 array-entered: 2, 2, 4, #NAME? (plain: #VALUE!, #VALUE!, 4, #NAME?).
        await Assert.That(Num(OnGrid("=SUM((A1:A3>Threshold)*1)"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUM((A1:A3>MyCell)*1)"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUM(Threshold*1)"))).IsEqualTo(4.0);
        await Assert.That(OnGrid("=SUM((Nope<>0)*1)")).IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task TwoDimensionalBareName_OnTheErrorFixture_KeepsTheEnginesScanOrder()
    {
        // LOAD-BEARING. Wide holds #N/A at B1 and #DIV/0! at A3. The engine's reference path scans a
        // rectangle COLUMN-major, so A3's #DIV/0! wins; the mini-CSE's RangeOperand is row-major, so on the
        // no-guard prototype every one of these went #DIV/0! → #N/A — a different error, silently, for a
        // shape that never asked to be an array. Each is pinned EQUAL to its literal twin, which is what
        // keeps this phase from moving the scan order by accident.
        // Aspose 26.6.0 answers #N/A in BOTH modes for both the name and the literal (row-major): that is a
        // measured, engine-wide divergence handed to Phase 11, NOT this slice's to fix.
        foreach (
            var (name, literal) in new[]
            {
                ("=SUM(Wide)", "=SUM(A1:B3)"),
                ("=SUBTOTAL(9,Wide)", "=SUBTOTAL(9,A1:B3)"),
                ("=SMALL(Wide,1)", "=SMALL(A1:B3,1)"),
                ("=MAX(Wide)", "=MAX(A1:B3)"),
                ("=AGGREGATE(9,4,Wide)", "=AGGREGATE(9,4,A1:B3)"),
            }
        )
        {
            var value = OnErrors(name);
            await Assert.That(value).IsEqualTo(OnErrors(literal));
            await Assert.That(value).IsEqualTo(ErrorValue.DivByZero);
        }
    }

    [Test]
    public async Task OpenRangeName_InAnArrayPosition_IsStillRefused_KnownDivergence()
    {
        // DELIBERATE DIVERGENCE, left as it is by this phase: MyCol = $A:$A, and the mini-CSE's cost guard
        // refuses a whole-column operand, so the name falls back to the scalar path and the truthy
        // reference value gives 1. Aspose 26.6.0: 0 PLAIN, 2 ARRAY-ENTERED (both columns quoted because
        // they disagree; MySheet implements the array-entered rule, so 2 is the value it eventually owes).
        await Assert.That(Num(OnGrid("=SUM((MyCol<>0)*1)"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task LetBoundName_InAnArrayPosition_IsUnchanged_KnownDivergence()
    {
        // DELIBERATE DIVERGENCE, left as it is by this phase: the consumer's argument is the Let node, which
        // Probe does not look inside, so the name arm never sees r. Phase 7's LET routing correction owns it.
        // Aspose 26.6.0: 2 PLAIN and 2 ARRAY-ENTERED (the two columns agree here).
        await Assert.That(Num(OnGrid("=SUM(LET(r,Rng,(r<>0)*1))"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task UnionName_InAnArrayPosition_IsUnchanged_KnownDivergence()
    {
        // DELIBERATE DIVERGENCE, left as it is by this phase, and it CONTRADICTS the sweep's B2 prediction:
        // a union name resolves to the Scalar outcome, the whole expression is then scalar-only, and the
        // consumer's scalar path evaluates the NameReference itself — so B2's resolved.Evaluate arm never
        // runs. The union ruling already owns this shape.
        // Aspose 26.6.0: #VALUE! PLAIN, 2 ARRAY-ENTERED (both quoted because they disagree). The literal
        // union twin answers #VALUE! here, so the name does NOT match its literal in this one row — pinned
        // as such rather than hidden.
        await Assert.That(Num(OnGrid("=SUM((UnN<>0)*1)"))).IsEqualTo(1.0);
        await Assert
            .That(OnGrid("=SUM(((Sheet1!$A$1:$A$2,Sheet1!$A$3:$A$3)<>0)*1)"))
            .IsEqualTo(ErrorValue.NotValue);
    }
}
