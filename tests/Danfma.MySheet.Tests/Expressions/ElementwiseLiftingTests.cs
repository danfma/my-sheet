using System.Reflection;
using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue; // TUnit also defines a StringValue

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 8 of <c>plans/structured-table-references-and-aggregate.md</c> — ELEMENTWISE LIFTING of unary
/// operators and pure-scalar built-ins inside the mini-CSE (<see cref="ArrayEvaluation"/>). Before this
/// phase the evaluator recognized only five array-producing shapes; a <c>UnaryOperation</c> or a scalar
/// function over a range fell to the opaque-scalar arm, was evaluated ONCE over a range (<c>#VALUE!</c>,
/// since a range has no scalar value) and broadcast that error into every element — or, worse, had it
/// silently CONSUMED by the enclosing node (a comparison, IFERROR's error arm, COUNT's non-numeric skip),
/// which is where the 0.0 answers pinned below came from.
///
/// <para>These are the phase's TDD pins, written BEFORE the engine change: every assertion whose comment says
/// "today" named the value measured on this tree at the time of writing, so the fix was a deliberate change
/// of answer rather than a silent one.</para>
///
/// <para>Oracles. Excel semantics per the Microsoft support pages "Guidelines and examples of array formulas"
/// (support.microsoft.com 7d94a64e-3ff3-4686-9372-ecfd5caa57c7) and "Implicit intersection operator: @"
/// (support.microsoft.com ce3be07b-0101-4450-a24e-c1c999be2b34), both fetched 2026-09-09; every golden below
/// was additionally MEASURED against the designated P0 oracle, Aspose.Cells 26.6.0, with the formula entered
/// BOTH plainly and CSE-entered (<c>SetArrayFormula</c>). The CSE column is the oracle for this file, because
/// that is the column the engine's array-consuming callers reproduce — the mini-CSE is entered only from a
/// consumed position (SUM/COUNT/SUMPRODUCT/SMALL/INDEX/AGGREGATE), never at the cell boundary, so a bare
/// <c>=LEN(A1:A3)</c> in a cell stays <c>#VALUE!</c> after this phase (Phase 7's business).</para>
/// </summary>
public class ElementwiseLiftingTests
{
    private const double Tolerance = 1e-9;

    // Numeric(): E6:E8 = 1,2,3 and F6:F8 = 10,20,30, and NOTHING else.
    //
    // The split from Textual() is deliberate and load-bearing. The phase file states one merged fixture that
    // also carries D7/D8/E9 text, but E7/E8/F7/F8 lie INSIDE D7:F9, so its own text goldens do not hold on
    // it: measured on the merged fixture, SUM(LEN(D7:F9)) is 13 and SUM(IF(LEN(D7:F9)>0,1,0)) is 7 (Aspose
    // CSE), not the 7 and 3 the plan asserts. Keeping the numbers away from D7:F9 is what makes both halves
    // of the plan true at once.
    private static object? OnNumeric(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["E6"] = new NumberValue(1);
        sheet["E7"] = new NumberValue(2);
        sheet["E8"] = new NumberValue(3);
        sheet["F6"] = new NumberValue(10);
        sheet["F7"] = new NumberValue(20);
        sheet["F8"] = new NumberValue(30);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Textual(): D7="abc", D8="def", E9=" " (ONE space), every other cell of D7:F9 empty. The lengths are
    // 3+3+0 / 0+0+1 / 0+0+0 = 7 raw and 3+3+0 / 0+0+0 / 0+0+0 = 6 after TRIM, so LEN and LEN(TRIM(…)) cannot
    // return the same number — the single space is the only element the two disagree on. Two spaces (as the
    // plan's prose writes it) would make the raw sum 8 and contradict the plan's own golden of 7.
    private static object? OnTextual(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["D7"] = new StringValue("abc");
        sheet["D8"] = new StringValue("def");
        sheet["E9"] = new StringValue(" ");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Broadcast(): Textual() plus E6:E8 = 1,2,3 — the ONE fixture where the merge Numeric()/Textual() avoid
    // is the point. E7 and E8 sit INSIDE the D7:F9 rectangle, so the rectangle's own elements are "abc", 2,
    // blank / "def", 3, blank / blank, " ", blank while the second argument is a 3x1 column of 1,2,3; each
    // row therefore cuts a DIFFERENT length, which is what makes a broadcast assertion distinguishable from
    // a fill (Textual()'s own goldens do not hold here — see the note on Numeric() above).
    private static object? OnBroadcast(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["D7"] = new StringValue("abc");
        sheet["D8"] = new StringValue("def");
        sheet["E9"] = new StringValue(" ");
        sheet["E6"] = new NumberValue(1);
        sheet["E7"] = new NumberValue(2);
        sheet["E8"] = new NumberValue(3);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Blank(): an empty sheet — A1:A3 are three BLANK elements, and Ghost! is a sheet that does not exist.
    private static object? OnBlank(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Errors(): C1 = =1/0, C2 = 5, C3 = 2 — deliberately NOT blank in C2/C3. With blanks there the lifted
    // answer for SUM(IFERROR(C1:C3,7)) is 7 and today's broken answer is ALSO 7 (the broadcast scalar
    // #VALUE! recovered once), so the assertion could never fail; measured both ways against Aspose CSE
    // (blank fixture 7, this fixture 14).
    private static object? OnErrors(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["C1"] = ExpressionParser.Parse("=1/0", sheet);
        sheet["C2"] = new NumberValue(5);
        sheet["C3"] = new NumberValue(2);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // MixedText(): B1 = 1, B2 = "x", B3 = 3 — one TEXT element inside a numeric column.
    private static object? OnMixedText(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["B1"] = new NumberValue(1);
        sheet["B2"] = new StringValue("x");
        sheet["B3"] = new NumberValue(3);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // RealDates(): D1:D3 = 2024-03-15/16/17 as DATE() serials, so the date-category pin measures the LIFT and
    // not this engine's epoch (see Date_Month_LiftsElementwise).
    private static object? OnRealDates(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["D1"] = ExpressionParser.Parse("=DATE(2024,3,15)", sheet);
        sheet["D2"] = ExpressionParser.Parse("=DATE(2024,3,16)", sheet);
        sheet["D3"] = ExpressionParser.Parse("=DATE(2024,3,17)", sheet);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // Lengths(): A1:A3 = 1, 22, 333 — three numbers whose TEXT LENGTHS are 1, 2, 3, all different. Any pin
    // that names a position (INDEX's n-th, AGGREGATE's k-th) is satisfiable by exactly one element, which a
    // fixture of equal-length numbers could never do.
    private static object? OnLengths(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(22);
        sheet["A3"] = new NumberValue(333);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // --- Acceptance: unary operators over an array ---

    [Test]
    public async Task Unary_DoubleNegation_CoercesTheComparisonArray()
    {
        // The `--` coercion idiom: E6:E8>1 is [FALSE,TRUE,TRUE] and each Negate lifts in turn, so the pair
        // turns the logicals into [0,1,1]. Today all three are #VALUE!: the outer Negate is an opaque scalar,
        // evaluated once over an array-shaped operand.
        await Assert.That(Num(OnNumeric("=SUMPRODUCT(--(E6:E8>1))"))).IsEqualTo(2.0);
        await Assert.That(Num(OnNumeric("=SUM(--(E6:E8>1))"))).IsEqualTo(2.0);

        // The single Negate keeps the SIGN, which is what distinguishes "lifted twice" from "lifted once":
        // -2 and 2 cannot both come from a coercion that lost the operator.
        await Assert.That(Num(OnNumeric("=SUM(-(E6:E8>1))"))).IsEqualTo(-2.0);
    }

    [Test]
    public async Task Unary_NegateAndPercent_OverARange_LiftElementwise()
    {
        // Negate and Percent are the only two unary operators this phase lifts. Today both are #VALUE!.
        await Assert.That(Num(OnNumeric("=SUM(-E6:E8)"))).IsEqualTo(-6.0);
        await Assert.That(Num(OnNumeric("=SUM(E6:E8%)"))).IsEqualTo(0.06).Within(Tolerance);
    }

    [Test]
    public async Task Unary_Plus_StaysAReferencePreservingNoOp()
    {
        // The NO-REGRESSION pin — the only pin here that is green before the fix AND asserts a real value
        // (the other two green ones, ShapeMismatch_… and TextElements_…, assert an error that today's opaque
        // scalar happens to produce for a different reason): unary `+`
        // is Excel's reference-preserving Lotus no-op (UnaryOperation routes Plus through
        // NamedReferences.CaptureValue), so +E6:E8 is still a REFERENCE and SUM reads its cells — 6 today,
        // and 6 after the phase, which must NOT lift Plus. Aspose agrees plainly and CSE-entered (6/6).
        await Assert.That(Num(OnNumeric("=SUM(+E6:E8)"))).IsEqualTo(6.0);
    }

    // --- Acceptance: one pin per function category ---

    [Test]
    public async Task Math_Abs_OverAnArrayOperand_LiftsElementwise()
    {
        // math category. E6:E8*-1 is already an eligible array (BinaryOperation); ABS is the node that must
        // lift over it. Today #VALUE!.
        await Assert.That(Num(OnNumeric("=SUM(ABS(E6:E8*-1))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Information_IsNumber_OverARange_LiftsElementwise()
    {
        // information category, and one of the three SILENT wrong answers this phase exists to remove: today
        // ISNUMBER(E6:E8) is an opaque scalar #VALUE!, the enclosing *1 consumes it, and the answer is 0.0 —
        // a plausible number no user sees as a failure. Excel (Aspose CSE): 3.
        await Assert.That(Num(OnNumeric("=SUM(ISNUMBER(E6:E8)*1)"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task Logical_IfError_OverARange_LiftsElementwise()
    {
        // logical category, second silent case: IFERROR's own error arm swallows the broadcast #VALUE! and
        // returns the fallback 0, so today the answer is 0.0 instead of 6.
        await Assert.That(Num(OnNumeric("=SUM(IFERROR(E6:E8,0))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Text_Len_OverARectangle_LiftsElementwise()
    {
        // text category over a 3x3 rectangle: 3+3+0 / 0+0+1 / 0+0+0 = 7. Today #VALUE!.
        await Assert.That(Num(OnTextual("=SUM(LEN(D7:F9))"))).IsEqualTo(7.0);

        // Third silent case, and the sharper of the two: COUNT skips non-numerics, so today's single opaque
        // #VALUE! counts 0 while the lifted 3x3 is nine numbers. 9 can only be the ELEMENT COUNT of the
        // rectangle — no cell content can produce it.
        await Assert.That(Num(OnTextual("=COUNT(LEN(D7:F9))"))).IsEqualTo(9.0);
    }

    [Test]
    public async Task Text_LenOverTrim_NestsTwoLifts()
    {
        // Two lifted nodes stacked: TRIM collapses E9's single space to "", so the sum drops from 7 to 6.
        // The pair 7/6 is what proves the inner node is lifted too rather than the outer one alone.
        await Assert.That(Num(OnTextual("=SUM(LEN(TRIM(D7:F9)))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task If_OverALiftedCondition_ZipsTheBranches()
    {
        // The already-eligible IfOperand consuming a LIFTED condition: LEN(D7:F9)>0 is [T,F,F / F,F,T / …],
        // so the count of non-empty cells is 3. Today #VALUE!: the condition itself was the opaque scalar.
        await Assert.That(Num(OnTextual("=SUM(IF(LEN(D7:F9)>0,1,0))"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task ProductionFormula_ShowHide_OverAnAnchoredRectangle()
    {
        // The end-to-end formula from the corpus that motivated the phase — anchored rectangle, TRIM, LEN, a
        // comparison, the `--` coercion, SUMPRODUCT and IF, all in one node. "Show" because three cells of
        // D7:F9 are non-blank and TRIM leaves two of them non-empty. Today #VALUE!.
        await Assert
            .That(OnTextual("=IF(SUMPRODUCT(--(LEN(TRIM($D$7:$F$9))>0))>0,\"Show\",\"Hide\")"))
            .IsEqualTo("Show");
    }

    [Test]
    public async Task Date_Month_LiftsElementwise()
    {
        // date category, pinned TWICE on purpose.
        //
        // (a) Over real dates — the form that measures the LIFT alone. MONTH of 2024-03-15/16/17 is 3,3,3, so
        // 9; Aspose CSE agrees. Today #VALUE!.
        await Assert.That(Num(OnRealDates("=SUM(MONTH(D1:D3))"))).IsEqualTo(9.0);

        // (b) Over the plan's own E6:E8 = 1,2,3. Phase 8 pinned 14 here — MONTH(1)=12, MONTH(2)=1,
        // MONTH(3)=1 under the OLE Automation epoch — precisely so that closing the epoch gap would have to
        // update this line deliberately. Phase 9 closed it: serial 1 is 1900-01-01, so MONTH is 1,1,1 and the
        // lifted sum is 3, which is what Aspose.Cells 26.6.0 answers (measured 2026-09-09, CSE column, where
        // this SUM-over-a-lifted-array shape lives; the PLAIN column is #VALUE! for both halves of this test).
        await Assert.That(Num(OnNumeric("=SUM(MONTH(E6:E8))"))).IsEqualTo(3.0);
    }

    // --- Edges: shape pairing and broadcast ---

    [Test]
    public async Task TwoArraysOfTheSameShape_PairElementwise()
    {
        // Two lifted arguments of the same shape zip position by position: ROUND(1,10), ROUND(2,20),
        // ROUND(3,30) = 1,2,3. Today #VALUE!. This is also the companion that keeps the mismatch pin below
        // from meaning "arrays simply do not work in a second argument".
        await Assert.That(Num(OnNumeric("=SUM(ROUND(E6:E8,F6:F8))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task ScalarArgument_Broadcasts_AndIsEvaluatedOnce()
    {
        // A scalar argument becomes a ScalarOperand — evaluated once at build time and broadcast — so the
        // array shape comes from the first argument alone. Today #VALUE!.
        await Assert.That(Num(OnNumeric("=SUM(ROUND(E6:E8,0))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task ShapeMismatch_BroadcastsTheColumnAcrossTheRectangle()
    {
        // A 3x3 first argument against a 3x1 second: the column repeats across every column of its row.
        //
        // FLIPPED by Phase 10 (was ShapeMismatch_FillsEveryElementWithValueError, pinning #VALUE!). Phase 8
        // pinned the engine's then-current dimension rule "with the divergence stated" precisely so that
        // closing the gap would be a deliberate update; the divergence — real Excel broadcasts a 3x1 column
        // across a 3x3 — is what Phase 10 closes, so the pin moves to the oracle's answers. The four
        // values this engine gave BEFORE the flip are kept here so the change of answer stays legible
        // (measured on 897affc, 2026-09-10): SUM(LEFT(D7:F9,E6:E8)) was #VALUE! on both fixtures,
        // SUM(LEN(LEFT(D7:F9,E6:E8))) was #VALUE!, INDEX(LEFT(D7:F9,E6:E8),2,1) was #VALUE! (it was #REF!
        // before Phase 8 lifted LEFT, which is the value the Phase 10 brief carried) and COUNT was already 0.
        //
        // Every golden below is Aspose.Cells 26.6.0, measured 2026-09-10, CSE column (plain entry answers
        // #VALUE! for the two SUMs and agrees on the INDEX).
        //
        // On Textual() E6:E8 are BLANK, so every element is LEFT(x, 0) = "" and the numeric fold is 0.
        await Assert.That(Num(OnTextual("=SUM(LEFT(D7:F9,E6:E8))"))).IsEqualTo(0.0);

        // A sum of nine empty strings is a weak assertion — it would also hold if the broadcast produced the
        // wrong element everywhere — so the same shape is pinned on Broadcast(), where the second argument
        // carries 1,2,3 and each row cuts a different length: "a","2","" / "de","3","" / ""," ","". LEN sums
        // to 6, nothing is numeric so COUNT is 0, and element (2,1) is the two-character "de".
        await Assert.That(Num(OnBroadcast("=SUM(LEN(LEFT(D7:F9,E6:E8)))"))).IsEqualTo(6.0);
        await Assert.That(Num(OnBroadcast("=COUNT(LEFT(D7:F9,E6:E8))"))).IsEqualTo(0.0);
        await Assert.That(OnBroadcast("=INDEX(LEFT(D7:F9,E6:E8),2,1)")).IsEqualTo("de");
    }

    // --- Edges: element kinds (blank, error, text, missing sheet) ---

    [Test]
    public async Task BlankElements_CoerceToZero()
    {
        // Three BLANK elements: blank coerces to 0 under Negate and LEN("") is 0, so both sums are 0.
        // Today #VALUE! for both.
        await Assert.That(Num(OnBlank("=SUM(-(A1:A3))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnBlank("=SUM(LEN(A1:A3))"))).IsEqualTo(0.0);

        // Anti-vacuity: 0.0 is also what a consumed error would sum to, so COUNT separates the two — three
        // zeros are three NUMBERS (Aspose CSE: 3), while today's single opaque #VALUE! counts 0.
        await Assert.That(Num(OnBlank("=COUNT(LEN(A1:A3))"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task ErrorElements_FirstInScanOrderWins()
    {
        // An error ELEMENT propagates through the fold, and the first one in row-major scan order wins.
        // E6:E8-2 is [-1,0,1], so the quotient is [-1,#DIV/0!,1] and ABS lifts over it: #DIV/0!, not the
        // #VALUE! of today's opaque scalar. Aspose CSE: #DIV/0!.
        await Assert.That(OnNumeric("=SUM(ABS(1/(E6:E8-2)))")).IsEqualTo(ErrorValue.DivByZero);

        // The same rule when the error comes from a CELL rather than from the arithmetic: C1 is #DIV/0!,
        // C2/C3 are 5 and 2, so LEN's elements are [#DIV/0!,1,1]. Today #VALUE!.
        await Assert.That(OnErrors("=SUM(LEN(C1:C3))")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Logical_IfError_RecoversPerElement()
    {
        // IFERROR lifted: the error element takes the fallback 7 and the two numbers pass through, so
        // 7+5+2 = 14. Today 7.0 — IFERROR recovered the ONE broadcast #VALUE! and SUM saw a single 7. The
        // fixture is what makes this fail: with C2/C3 blank both answers would be 7 (measured both ways).
        await Assert.That(Num(OnErrors("=SUM(IFERROR(C1:C3,7))"))).IsEqualTo(14.0);
    }

    [Test]
    public async Task TextElements_UnderANumericLift_AreValueErrors()
    {
        // B2 = "x" cannot coerce to a number, so its element is #VALUE! and the fold reports it. Aspose CSE
        // agrees (#VALUE!), and so does today's tree — for the WRONG reason (the opaque scalar), which is why
        // this pair is a regression pin for the element-kind rule rather than a failing acceptance case.
        await Assert.That(OnMixedText("=SUM(-B1:B3)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnMixedText("=SUM(ABS(B1:B3))")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task MissingSheet_IsAReferenceError_NotAValueError()
    {
        // The element reader (RangeOperand.At → Workbook.GetCellValueDense) owns the missing-sheet rule, so a
        // lifted LEN over a range on a sheet that does not exist reports #REF! — Excel's answer (Aspose CSE:
        // #REF!) and strictly better than today's #VALUE!.
        //
        // It also contrasts with the documented SUM(ROW(Ghost!A1:A3)) = 6 divergence pinned in
        // MiniCseConsumerTests, which survives because PositionNumbersOperand never touches a cell.
        await Assert.That(OnBlank("=SUM(LEN(Ghost!A1:A3))")).IsEqualTo(ErrorValue.Reference);
    }

    // --- The three known divergences this phase LEAVES OPEN, pinned as gaps ---
    //
    // docs/workbook-and-expressions.md's "Known divergences" list promises that every entry on it is pinned by
    // a test as a GAP rather than asserted as Excel's rule, so closing one is always a deliberate edit. These
    // three tests are that promise for the three entries this phase added to the list. Every oracle number in
    // them is Aspose.Cells 26.6.0, CSE-entered (SetArrayFormula), measured 2026-09-09; plain entry answers
    // #VALUE! for all of them on both engines except where noted, so the CSE column is the one that diverges.

    [Test]
    public async Task LiftedCall_UnderAnOpaqueUnaryPlus_IsNotLifted_KnownDivergence()
    {
        // Unary '+' is Excel's reference-preserving no-op, and MySheet keeps the whole '+' expression opaque
        // (UnaryOperation.Evaluate captures the operand as a reference VALUE; the mini-CSE's builder excludes
        // Plus). That hides what is INSIDE it, so a lifted call under a '+' is not lifted. The oracle lifts
        // it: SUM(+LEN(A1:A3)) = 6 and SUM(LEN(+A1:A3)) = 6, both against the #VALUE! pinned here. This is
        // the sibling of the already-documented SUM(-(+A1:A3)) = -6 gap — the same opaque '+', with a lifted
        // FUNCTION inside it instead of a unary operator.
        await Assert.That(OnLengths("=SUM(+LEN(A1:A3))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnLengths("=SUM(LEN(+A1:A3))")).IsEqualTo(ErrorValue.NotValue);

        // The control, and what keeps the gap narrow: without the '+' the same shape lifts (1 + 2 + 3), and
        // the '+' over a bare range still reads the cells on the ordinary range path (1 + 22 + 333). So the
        // gap is exactly "a lifted shape wrapped in '+'", not "'+' loses the cells".
        await Assert.That(Num(OnLengths("=SUM(LEN(A1:A3))"))).IsEqualTo(6.0);
        await Assert.That(Num(OnLengths("=SUM(+A1:A3)"))).IsEqualTo(356.0);
    }

    // MyName = Sheet1!$A$1:$A$3 over the Lengths fixture, so the name denotes the same 1, 22, 333 the tests
    // above write out literally — the only difference between the two halves of the pin is the NAME.
    private static object? OnNamedLengths(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(22);
        sheet["A3"] = new NumberValue(333);
        workbook.DefineName("MyName", "Sheet1!$A$1:$A$3");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task LiftedShapes_OverADefinedName_AreNotLifted_KnownDivergence()
    {
        // A defined name is captured as a reference VALUE, so it reaches the mini-CSE as an opaque scalar
        // unless the consuming shape resolves the name itself. ROW/COLUMN do (Phase 7 gave them that path);
        // nothing else does, so EVERY other array shape over a name is a gap. Oracle, against the answers
        // pinned here: SUM(LEN(MyName)) = 6, SUM(-MyName) = -356, SUM(MyName%) = 3.56, SUM(MyName*2) = 712.
        await Assert.That(OnNamedLengths("=SUM(LEN(MyName))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnNamedLengths("=SUM(-MyName)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnNamedLengths("=SUM(MyName%)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnNamedLengths("=SUM(MyName*2)")).IsEqualTo(ErrorValue.NotValue);

        // The COMPARISON shapes are the load-bearing half of this pin, because they are SILENT rather than
        // errors: the name collapses to its first cell, 1 > 1 is FALSE, and IF's else branch answers 1 — a
        // plausible number where the oracle answers 2 (22 and 333 both exceed 1). A reader who only saw the
        // #VALUE!s above would think the gap always announces itself.
        await Assert.That(Num(OnNamedLengths("=SUM(IF(MyName>1,1,0))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnNamedLengths("=SUMPRODUCT(--(MyName>1))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnNamedLengths("=SUM((MyName>1)*1)"))).IsEqualTo(1.0);

        // The controls: reading the name is unaffected (356 on both engines), and ROW over it IS lifted
        // (6 on both), which is what makes this a gap about the LIFTED shapes and not about names.
        await Assert.That(Num(OnNamedLengths("=SUM(MyName)"))).IsEqualTo(356.0);
        await Assert.That(Num(OnNamedLengths("=SUM(ROW(MyName))"))).IsEqualTo(6.0);
    }

    // A1:A3 = 1,2,3 and B1:B3 = 10,20,30 — the fixture the per-slot measurements were taken on. The values
    // are small and ascending so that a lifted answer and a first-element answer are always different
    // numbers (MATCH 6 against 1, VLOOKUP 60 against 10, LARGE 6 against 3).
    private static object? OnSlots(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        sheet["B1"] = new NumberValue(10);
        sheet["B2"] = new NumberValue(20);
        sheet["B3"] = new NumberValue(30);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task AConsumesFunction_IsNotLiftedOverItsScalarSlots_KnownDivergence()
    {
        // The classification is per FUNCTION, not per slot: a Consumes entry is never lifted, not even over
        // the slots that take a scalar. Excel's rule is the other one — it consumes the range in the slot
        // that takes one and repeats the WHOLE CALL per element of a rectangle handed to any other slot. The
        // oracle number follows each assertion; the assertion is what this engine answers today.
        await Assert.That(OnSlots("=SUM(MATCH(A1:A3,A1:A3,0))")).IsEqualTo(ErrorValue.NotAvailable); // 6
        await Assert
            .That(OnSlots("=SUM(VLOOKUP(A1:A3,A1:B3,2,FALSE))"))
            .IsEqualTo(ErrorValue.NotValue); // 60
        await Assert.That(OnSlots("=SUM(CHOOSE(A1:A3,10,20,30))")).IsEqualTo(ErrorValue.NotValue); // 60
        await Assert.That(OnSlots("=SUM(LARGE(A1:A3,A1:A3))")).IsEqualTo(ErrorValue.NotValue); // 6
        await Assert.That(OnSlots("=SUM(INDEX(B1:B3,A1:A3))")).IsEqualTo(ErrorValue.NotValue); // 60
        await Assert.That(OnSlots("=SUM(RANK(A1:A3,A1:A3))")).IsEqualTo(ErrorValue.NotValue); // 6
        await Assert.That(OnSlots("=SUM(WORKDAY(A1:A3,1))")).IsEqualTo(ErrorValue.NotValue); // 9
        await Assert
            .That(OnSlots("=SUM(NETWORKDAYS.INTL(A1:A3,4))"))
            .IsEqualTo(ErrorValue.NotValue); // 9
        await Assert.That(OnSlots("=SUM(NPV(A1:A3/10,10,20,30))")).IsEqualTo(ErrorValue.NotValue); // 120.92
        await Assert.That(OnSlots("=SUM(RANDBETWEEN(A1:A3,A1:A3))")).IsEqualTo(ErrorValue.NotValue); // 6

        // Plain NETWORKDAYS is the one member of the family the ORACLE does not lift either: there
        // SUM(NETWORKDAYS(A1:A3,B1:B3)) is 8, which is NETWORKDAYS(A1,B1) alone — an implicit intersection to
        // the first element, not a per-element lift (a lift would be 8 + 14 + 20). Named so nobody "fixes"
        // this engine towards 42.
        await Assert.That(OnSlots("=SUM(NETWORKDAYS(A1:A3,B1:B3))")).IsEqualTo(ErrorValue.NotValue); // 8

        // The two SILENT answers, and the reason this pin is not a list of #VALUE!s. COUNTIF's collapsed
        // argument matches no criterion, so the scan comes back empty and the count is 0 where the oracle
        // lifts the criteria slot and answers 3; TYPE is handed the #VALUE! of a range in a scalar slot and
        // reports its TYPE CODE, 16, where the oracle answers 1+1+1 = 3. Both are plausible numbers.
        await Assert.That(Num(OnSlots("=SUM(COUNTIF(A1:A3,A1:A3))"))).IsEqualTo(0.0); // 3
        await Assert.That(Num(OnSlots("=SUM(TYPE(A1:A3))"))).IsEqualTo(16.0); // 3
    }

    // --- Item 9: the classification guard (the regression defence for the next contributor) ---

    // THREE rectangles that differ in POSITION, in SHAPE and in CONTENTS (and in the KIND of their
    // contents), plus the rectangle the probe puts in the slots that are NOT under test. Handing a
    // function one rectangle and then another is the range-awareness probe that produced the 180/126 split,
    // re-run on every build instead of trusting a frozen list — see EveryPositionProbe for what it can and
    // cannot see.
    private static (Workbook Workbook, Sheet Sheet) ThreeRectangles()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // R1 = A1:A3 — a 3x1 column of small ascending numbers.
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);

        // R2 = C4:C6 — a 3x1 column elsewhere, holding a large number, a text, and a BLANK (C6 is
        // deliberately left empty — a third element kind the other two rectangles do not have).
        sheet["C4"] = new NumberValue(100);
        sheet["C5"] = new StringValue("x");

        // R3 = E1:G1 — a 1x3 ROW, so a body that reads the SHAPE rather than the contents (ROWS, COLUMNS,
        // MATCH's match_type over a row) tells it apart from the two columns.
        sheet["E1"] = new NumberValue(-7);
        sheet["F1"] = new BooleanValue(true);
        sheet["G1"] = new StringValue("zz");

        // H1:H3 — the RANGE filler, for the slots the sweep is not currently probing. It is deliberately
        // THREE cells, the same element count as each rectangle, so that a body needing a SECOND population
        // of matching length (CORREL, SUMXMY2, SUMIFS's criteria range) gets one and can therefore answer
        // differently for two rectangles instead of erroring identically on all of them.
        sheet["H1"] = new NumberValue(4);
        sheet["H2"] = new NumberValue(5);
        sheet["H3"] = new NumberValue(6);

        return (workbook, sheet);
    }

    // The three rectangles, in the order the probe compares them.
    private static readonly string[] Rectangles = ["A1:A3", "C4:C6", "E1:G1"];

    // What the probe puts in every slot that is NOT holding the rectangle. Four kinds, because a body can be
    // range-aware in one slot only while the OTHER slots decide whether it gets that far: a number, a text, a
    // logical, and a rectangle of its own (which is what un-blinds the functions whose range slot is the one
    // the sweep is currently filling — VLOOKUP's table, NETWORKDAYS's holidays, CORREL's second array).
    private static readonly string[] Fillers = ["1", "\"a\"", "TRUE", "H1:H3"];

    // THE probe, shared by the oracle test (over the Elementwise half of the registry) and by the
    // completeness assertion (over the Consumes half), so the two can never disagree about what "blind"
    // means. It SWEEPS: for every arity from MinArgs to MinArgs+3 (capped at MaxArgs), every argument
    // POSITION in that arity, and every filler for the remaining slots, it hands the entry each of the three
    // rectangles in turn and evaluates on the ORDINARY scalar path, which never enters the mini-CSE. The
    // first call whose answers are not all equal is returned as a diagnostic; `null` means the entry is BLIND
    // to the whole sweep.
    //
    // What it sees: any body that READS a rectangle handed to ANY of its slots (1,2,3 against 100,"x",blank
    // against -7,TRUE,"zz"), and any body that reads a rectangle's SHAPE (3x1 against 1x3). What it still
    // CANNOT see: a body that answers the same thing for every rectangle — a reference test (ISREF,
    // ISFORMULA), a type code, a fold over a whole population that errors identically on all three, or a
    // slot whose range-awareness needs a second population the fillers cannot supply. Those are blind here
    // and must be NAMED — which is what TheShapeAndPositionAndCriteriaFamilies_StayConsumes does, and what
    // NoConsumesEntry_IsBlindToTheProbe_AndNamedNowhere proves is exhaustive.
    private static string? EveryPositionProbe(
        FunctionRegistry.RegistryEntry entry,
        Workbook workbook,
        Sheet sheet
    )
    {
        var top = Math.Min(entry.MaxArgs, entry.MinArgs + 3);

        for (var arity = Math.Max(1, entry.MinArgs); arity <= top; arity++)
        {
            for (var position = 0; position < arity; position++)
            {
                foreach (var filler in Fillers)
                {
                    var answers = Rectangles
                        .Select(rectangle => Answer(Call(rectangle, arity, position, filler)))
                        .ToArray();

                    for (var other = 1; other < answers.Length; other++)
                    {
                        if (Equals(answers[0], answers[other]))
                        {
                            continue;
                        }

                        return $"{Call(Rectangles[0], arity, position, filler)} → {answers[0]}, "
                            + $"{Call(Rectangles[other], arity, position, filler)} → {answers[other]}";
                    }
                }
            }
        }

        return null;

        string Call(string rectangle, int arity, int position, string filler)
        {
            var slots = new string[arity];

            for (var slot = 0; slot < arity; slot++)
            {
                slots[slot] = slot == position ? rectangle : filler;
            }

            return $"{entry.Name}({string.Join(',', slots)})";
        }

        object? Answer(string body)
        {
            try
            {
                return ExpressionParser.ParseFormulaBody(body, sheet).Evaluate(workbook).AsObject();
            }
            catch (Exception exception)
            {
                // A throw is an answer too: every rectangle must still reach the SAME one.
                return $"{exception.GetType().Name}: {exception.Message}";
            }
        }
    }

    [Test]
    public async Task EveryElementwiseEntry_IsBlindToTheRangeItIsHanded()
    {
        // THE guard. For every entry the registry flags Elementwise, SWEEP the probe over it — every arity
        // from MinArgs to MinArgs+3, every argument position in that arity, each of four fillers in the
        // remaining slots — handing it three different rectangles in turn on the ORDINARY scalar path
        // (Parse(...).Evaluate, which never enters the mini-CSE) and require every call to answer IDENTICALLY
        // for all three. A pure-scalar body cannot tell the rectangles apart: a range has no scalar value, so
        // it answers the same #VALUE! (or the same constant made from the other arguments) for each. A
        // RANGE-AWARE body reads the cells (1,2,3 against 100,"x",blank against -7,TRUE,"zz") or their shape
        // (3x1 against 1x3) and therefore answers differently — which is exactly the mis-flag this test
        // exists to catch, and the mis-flag that would otherwise ship as a SILENT wrong number (the lift
        // would hand the function one element of the rectangle it was meant to consume whole; see the three
        // measured silent cases in this file's class comment).
        //
        // KNOWN LIMITATION, and the reason the named list below exists. The sweep sees a body that reads a
        // rectangle in ANY slot, but not one that answers the same thing for every rectangle: a reference
        // test, a type code, or a fold that errors identically on all three. Measured, 21 of the 126 Consumes
        // entries are still blind, so this is the CHEAP, always-on half of the derivation, not the derivation
        // itself, and TheShapeAndPositionAndCriteriaFamilies_StayConsumes is its complement: what the probe
        // cannot see must be named by hand.
        //
        // Verified by mutation, three times. Flagging SUM `Elementwise` fails with
        // "SUM(A1:A3) → 6, SUM(C4:C6) → 100"; flagging COUNT — a range-aware function the explicit list
        // below does NOT name — fails with "COUNT(A1:A3) → 3, COUNT(C4:C6) → 1"; and flagging HLOOKUP, whose
        // table sits in a LATER slot and which the old position-0 probe could not see at all, now fails with
        // "HLOOKUP: HLOOKUP(1,A1:A3,1) → 1, HLOOKUP(1,C4:C6,1) → #N/A". The last two are the point: the
        // sweep catches a mis-flag nobody remembered to enumerate — the HLOOKUP mutation, made the way a
        // contributor would (flip the factory, delete the [Arguments("HLOOKUP")] row below, bump the four
        // count constants), left the whole suite GREEN at 1485/0 under the position-0 probe.
        //
        // Item 9(a) rides along: arguments can only be built when the entry takes at least one, and a
        // zero-argument entry has nothing to lift (FunctionRegistryClassificationTests pins the same rule
        // from the registry side).
        //
        // NOT oracle-verified: sixteen of the 180 names this loop walks are functions Aspose.Cells 26.6.0
        // does not implement (#NAME? for every call), so their Elementwise flag is inferred from the node
        // bodies and from this sweep rather than measured. FunctionRegistryClassificationTests
        // .TheElementwiseRoster names all sixteen.
        var (workbook, sheet) = ThreeRectangles();
        var lifted = new List<string>();
        var offenders = new List<string>();

        foreach (var entry in FunctionRegistry.ByName.Values)
        {
            if (entry.Lifting is not ArrayLifting.Elementwise)
            {
                continue;
            }

            lifted.Add(entry.Name);

            if (entry.MaxArgs is 0)
            {
                offenders.Add($"{entry.Name}: Elementwise with MaxArgs 0 — nothing to lift over");
                continue;
            }

            if (EveryPositionProbe(entry, workbook, sheet) is { } discriminated)
            {
                offenders.Add($"{entry.Name}: {discriminated}");
            }
        }

        await Assert.That(offenders).IsEmpty();
        // Anti-vacuity: the loop must actually have walked the whole lifted half of the registry.
        await Assert.That(lifted.Count).IsEqualTo(180);
    }

    [Test]
    public async Task EveryElementwiseEntryWithAnOptionalSlot_KeepsTheOmittedSlotUnderTheLift()
    {
        // Verifier correction M2, and the guard that the scalar-blindness sweep above cannot be: it compares
        // the LIFTED element against the SCALAR answer for the same cell, so it fails whenever the lift
        // changes what the body sees. The shape it drives is the one B1 fixed — an OMITTED optional slot,
        // which ten of the lifted built-ins (FIXED, DOLLAR, NUMBERVALUE, TEXTBEFORE/TEXTAFTER, VALUETOTEXT,
        // the REGEX family, ADDRESS) detect by pattern-matching the parser's literal BlankValue node. A
        // scratch slot there silently turns "omitted" into "blank, coerced to 0"; measured, that costs
        // FIXED(A1:A3,,TRUE) its two decimals.
        //
        // Verified by mutation: dropping B1's `arguments[j] is BlankValue ? arguments[j] : …` guard in
        // LiftedFunctionOperand fails this test on ADDRESS, DOLLAR, TEXTAFTER, NUMBERVALUE and the rest of
        // the ten — e.g. "DOLLAR(A1:A3,)[0] → $1, scalar DOLLAR(A1,) → $1.00".
        //
        // The arity is the smallest one that HAS a slot to omit plus, where the entry allows it, one slot
        // after it — `F(A1:A3, 1, …, , 1)` — because a trailing omission is indistinguishable from simply
        // writing fewer arguments. The filler 1 makes several of these answer an error; that is fine and
        // still discriminating, since the assertion is lifted == scalar, not lifted == some golden.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(22);
        sheet["A3"] = new NumberValue(333);

        var context = new EvaluationContext(workbook, sheet.Name);
        var covered = new List<string>();
        var offenders = new List<string>();

        foreach (var entry in FunctionRegistry.ByName.Values)
        {
            if (entry.Lifting is not ArrayLifting.Elementwise || entry.MaxArgs <= entry.MinArgs)
            {
                continue;
            }

            covered.Add(entry.Name);

            // MinArgs + 2 keeps a slot AFTER the omitted one wherever the entry has room for it; the cap is
            // MaxArgs (int.MaxValue for the variadic IFS/SWITCH, so the cap never binds there).
            var arity = Math.Min(entry.MaxArgs, entry.MinArgs + 2);
            var omitted = arity >= 3 ? arity - 2 : arity - 1;

            string Call(string first)
            {
                var slots = new string[arity];
                slots[0] = first;

                for (var i = 1; i < arity; i++)
                {
                    slots[i] = i == omitted ? string.Empty : "1";
                }

                return $"{entry.Name}({string.Join(',', slots)})";
            }

            if (
                !ArrayEvaluation.TryEvaluate(
                    ExpressionParser.ParseFormulaBody(Call("A1:A3"), sheet),
                    context,
                    out var result
                )
            )
            {
                offenders.Add($"{Call("A1:A3")}: not array-eligible — the lift did not fire");
                continue;
            }

            for (var element = 0; element < 3; element++)
            {
                var lifted = result.Values[element].AsObject();
                var scalar = ExpressionParser
                    .ParseFormulaBody(Call($"A{element + 1}"), sheet)
                    .Evaluate(context)
                    .AsObject();

                if (!Equals(lifted, scalar))
                {
                    offenders.Add(
                        $"{Call("A1:A3")}[{element}] → {lifted}, scalar {Call($"A{element + 1}")} → {scalar}"
                    );
                }
            }
        }

        await Assert.That(offenders).IsEmpty();
        // Anti-vacuity: 65 of the 180 lifted entries have at least one optional slot.
        await Assert.That(covered.Count).IsEqualTo(65);
    }

    [Test]
    [Arguments("SUM")]
    [Arguments("SUMPRODUCT")]
    [Arguments("INDEX")]
    [Arguments("ROW")]
    [Arguments("COLUMN")]
    [Arguments("ROWS")]
    [Arguments("COLUMNS")]
    [Arguments("AREAS")]
    [Arguments("OFFSET")]
    [Arguments("INDIRECT")]
    [Arguments("ISREF")]
    [Arguments("TYPE")]
    [Arguments("MATCH")]
    [Arguments("VLOOKUP")]
    // The lookup family in full. These were the position-0 probe's blindest spot — HLOOKUP, VLOOKUP,
    // XLOOKUP, LOOKUP, MATCH and XMATCH all take their table/lookup_array in a LATER slot, and CHOOSE's first
    // slot is the index — which is what motivated the sweep over every position: measured, mutating
    // Entry<HLookup> to Elementwise<HLookup> left the position-0 oracle's offender list EMPTY, while the
    // sweep now fails it with a diagnostic. They stay named all the same, because a name is the cheapest
    // documentation of WHY each one consumes.
    [Arguments("HLOOKUP")]
    [Arguments("XLOOKUP")]
    [Arguments("XMATCH")]
    [Arguments("LOOKUP")]
    [Arguments("CHOOSE")]
    // The reference-TAKING siblings of ISREF/TYPE/ROW/COLUMN. Their argument is a reference whose CELLS they
    // never read, so EVERY rectangle looks alike to them and the sweep is blind to them for good: a lift
    // would hand them one element and lose the reference entirely.
    [Arguments("ISFORMULA")]
    [Arguments("FORMULATEXT")]
    [Arguments("SHEET")]
    [Arguments("SUBTOTAL")]
    [Arguments("AGGREGATE")]
    [Arguments("IF")]
    [Arguments("LET")]
    [Arguments("RANDBETWEEN")]
    // The whole criteria family — every *IF/*IFS built-in. CriteriaScan.Open reads the range itself, so a
    // lift would hand it one element and answer from that alone.
    [Arguments("SUMIF")]
    [Arguments("SUMIFS")]
    [Arguments("COUNTIF")]
    [Arguments("COUNTIFS")]
    [Arguments("AVERAGEIF")]
    [Arguments("AVERAGEIFS")]
    [Arguments("MAXIFS")]
    [Arguments("MINIFS")]
    // The logical folds. AND/OR/XOR reduce every cell of every argument to one truth value, so a lift would
    // answer from a single element and silently return the wrong verdict; TEXTJOIN concatenates them.
    [Arguments("AND")]
    [Arguments("OR")]
    [Arguments("XOR")]
    [Arguments("TEXTJOIN")]
    // The cash-flow functions: the SERIES is the argument. NPV/XNPV/IRR/MIRR/XIRR/FVSCHEDULE/SERIESSUM each
    // walk a whole vector of values (and XNPV/XIRR a parallel vector of dates), so one element is not a
    // smaller answer — it is a different function.
    [Arguments("NPV")]
    [Arguments("XNPV")]
    [Arguments("IRR")]
    [Arguments("MIRR")]
    [Arguments("XIRR")]
    [Arguments("FVSCHEDULE")]
    [Arguments("SERIESSUM")]
    // The population statistics: RANK*/MODE*/KURT/PROB/INTERCEPT are defined over the whole population (a
    // rank needs every other value to compare against, a mode needs the frequencies), which is precisely what
    // an elementwise lift destroys.
    [Arguments("RANK")]
    [Arguments("RANK.EQ")]
    [Arguments("RANK.AVG")]
    [Arguments("MODE")]
    [Arguments("MODE.SNGL")]
    [Arguments("KURT")]
    [Arguments("PROB")]
    [Arguments("INTERCEPT")]
    // The holiday-aware date functions: their optional `holidays` argument is a RANGE, in a later slot — the
    // same shape that hid the lookup family above from the old position-0 probe.
    [Arguments("NETWORKDAYS")]
    [Arguments("NETWORKDAYS.INTL")]
    [Arguments("WORKDAY")]
    [Arguments("WORKDAY.INTL")]
    public async Task TheShapeAndPositionAndCriteriaFamilies_StayConsumes(string name)
    {
        // The explicit half of the guard, and the COMPLEMENT to the executable oracle above. The sweep sees
        // most of these now; the ones it cannot see (the shapes, the reference tests, the type code, the
        // whole-population folds — the 21 pinned by name in
        // NoConsumesEntry_IsBlindToTheProbe_AndNamedNowhere) answer the SAME thing for every rectangle in
        // every slot, so scalar-blindness holds for them while they are still range-aware, and being named
        // here is the only defence they have. IF and RANDBETWEEN are design exclusions (IF owns a dedicated
        // operand arm; lifting a volatile would draw once per element) and LET binds names to whole
        // sub-expressions.
        //
        // Deliberately NOT in this list: the logical IFS and DATEDIF, whose names end in IF/IFS but which are
        // pure-scalar and therefore correctly Elementwise.
        await Assert.That(FunctionRegistry.ByName[name].Lifting).IsEqualTo(ArrayLifting.Consumes);
    }

    // --- Item 11: INDEX's widened rejection, and the array path it now falls into ---

    [Test]
    public async Task Index_OverALiftedArgument_TakesTheArrayPath()
    {
        // Index.TryResolveReference returns false — falling back to normal evaluation — when its first
        // argument is not a Reference but IS array-eligible. This phase widens "array-eligible", so
        // INDEX(LEN(…), n) now takes the ARRAY path where it previously took the reference path and failed.
        // The pin makes that widening intentional and visible; it is the ONE integration point where this
        // phase changes an existing decision rather than adding a new one.
        //
        // Aspose.Cells 26.6.0, measured 2026-09-09 (plain AND CSE-entered, identical): INDEX(LEN(D7:F9),1)
        // = 3 and ROWS(INDEX(LEN(D7:F9),1)) = 1.
        await Assert.That(Num(OnTextual("=INDEX(LEN(D7:F9),1)"))).IsEqualTo(3.0);

        // A lifted INDEX result is a VALUE, not a reference — it has no cell address, and INDEX deliberately
        // does not invent one. A scalar counts as 1x1 (Phase 1 item 2), so ROWS of it is 1.
        await Assert.That(Num(OnTextual("=ROWS(INDEX(LEN(D7:F9),1))"))).IsEqualTo(1.0);

        // The positional half, on the fixture whose element lengths are all different, so n selects. Aspose
        // (26.6.0, 2026-09-09, plain and CSE): 1 / 2 / 3, and #REF! past the end.
        await Assert.That(Num(OnLengths("=INDEX(LEN(A1:A3),1)"))).IsEqualTo(1.0);
        await Assert.That(Num(OnLengths("=INDEX(LEN(A1:A3),2)"))).IsEqualTo(2.0);
        await Assert.That(Num(OnLengths("=INDEX(LEN(A1:A3),3)"))).IsEqualTo(3.0);
        await Assert.That(OnLengths("=INDEX(LEN(A1:A3),4)")).IsEqualTo(ErrorValue.Reference);

        // The unary half of the lift reaches INDEX the same way. Aspose: -22.
        await Assert.That(Num(OnLengths("=INDEX(-A1:A3,2)"))).IsEqualTo(-22.0);
    }

    // --- The AGGREGATE/SUBTOTAL reference slot: Phase 2's rule meets the widened eligibility ---

    [Test]
    public async Task Subtotal_And_Aggregate_ReferenceSlot_RefuseALiftedArgument()
    {
        // Phase 2 gave AggregateCodes.Feed Excel's REF-SLOT rule: function_num 1-13 of SUBTOTAL and
        // AGGREGATE take a REFERENCE, so a non-reference array-eligible argument is #VALUE!. Making LEN
        // Elementwise brings LEN(A1:A3) into "array-eligible", so these two forms flip from a broadcast
        // scalar error to a deliberate, Excel-correct refusal — the SAME #VALUE!, now for the right reason.
        // Aspose.Cells 26.6.0, measured 2026-09-09 (plain and CSE-entered, identical).
        await Assert.That(OnLengths("=SUBTOTAL(9,LEN(A1:A3))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnLengths("=AGGREGATE(9,4,LEN(A1:A3))")).IsEqualTo(ErrorValue.NotValue);

        // The unary lift lands in the same slot and is refused identically (Aspose: #VALUE! for both).
        await Assert.That(OnLengths("=SUBTOTAL(9,-A1:A3)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnLengths("=AGGREGATE(9,4,-A1:A3)")).IsEqualTo(ErrorValue.NotValue);

        // AGGREGATE's ARRAY form (14-19) takes an array, so there the SAME lifted argument is streamed —
        // which is what makes the four refusals above a rule about the SLOT rather than about the lift.
        // Aspose: SMALL k=1 → 1, k=3 → 3, LARGE k=1 → 3.
        await Assert.That(Num(OnLengths("=AGGREGATE(15,6,LEN(A1:A3),1)"))).IsEqualTo(1.0);
        await Assert.That(Num(OnLengths("=AGGREGATE(15,6,LEN(A1:A3),3)"))).IsEqualTo(3.0);
        await Assert.That(Num(OnLengths("=AGGREGATE(14,4,LEN(A1:A3),1)"))).IsEqualTo(3.0);
    }

    // --- Item 12 (c) and the end-to-end integration through a real cell ---

    [Test]
    public async Task LiftedFormula_ThroughARealCell_AndTheDirtyGraph()
    {
        // The consumers are reached through Workbook.EvaluateCell/GetCellValue in production, not through
        // Parse(...).Evaluate as every other case in this file. This drives one lifted formula end to end:
        // stored in a real cell, read back through the host API, then INVALIDATED through the dirty graph by
        // an edit to a cell of the range the lift reads.
        //
        // The dirty half is the load-bearing one. DependencyExtractor reaches a function's arguments through
        // FormulaWriter.Call, so the range inside LEN is a real RangeDep and B1 lands in the dirty cone of an
        // A3 edit (DependencyExtractorTests pins the scan itself). If the lift had hidden its argument from
        // the extractor, B1 would answer the stale 6 here.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(22);
        sheet["A3"] = new NumberValue(333);
        sheet["B1"] = ExpressionParser.Parse("=SUM(LEN(A1:A3))", sheet);

        var engine = DirtyEngine.Build(workbook);

        // 1 + 2 + 3 = 6 (Aspose.Cells 26.6.0, CSE-entered, measured 2026-09-09).
        await Assert.That(workbook.GetCellValue("Sheet1", "B1").AsObject()).IsEqualTo(6.0);

        sheet["A3"] = new NumberValue(1234);
        var dirty = engine.CalculateDirty([new CellDep("Sheet1", 1, 3)]);

        // B1 is in the cone — the lifted node's range dependency reached the graph.
        await Assert.That(dirty!.Contains(new CellDep("Sheet1", 2, 1))).IsTrue();

        // … and the evicted cell recomputes to 1 + 2 + 4 = 7 (Aspose 26.6.0, CSE-entered, 2026-09-09).
        await Assert.That(workbook.GetCellValue("Sheet1", "B1").AsObject()).IsEqualTo(7.0);
    }

    [Test]
    public async Task AnchoredMaster_OverALiftedFormula_StaysFullyAnchored()
    {
        // Item 12 (c). AnchoredFormulaSupport decides whether a shared-formula group may share ONE anchored
        // master tree. Its `Function function =>` arm accepts a function when every argument is anchored,
        // reaching the arguments through the same registry accessor (FormulaWriter.Call) the writer and the
        // dependency extractor use — and this phase adds no node type, so there is nothing new for it to
        // learn. The pin says so out loud, because a lifted formula is exactly the shape a reader would
        // expect to have broken it.
        var sheet = new Sheet { Name = "Sheet1" };
        var tokens = ExpressionParser.TokenizeFormulaBody("SUM(LEN($A$1:$A$3))");
        var master = ExpressionParser.ParseAnchoredMasterBody(tokens, sheet);

        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(master)).IsTrue();

        // The unary half too, and a control that still falls back (an open range is not anchorable).
        var unary = ExpressionParser.ParseAnchoredMasterBody(
            ExpressionParser.TokenizeFormulaBody("SUM(-($A$1:$A$3>1))"),
            sheet
        );
        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(unary)).IsTrue();

        var openRange = ExpressionParser.ParseAnchoredMasterBody(
            ExpressionParser.TokenizeFormulaBody("SUM(LEN($A:$A))"),
            sheet
        );
        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(openRange)).IsFalse();
    }

    [Test]
    public async Task NoConsumesEntry_IsBlindToTheProbe_AndNamedNowhere()
    {
        // THE CLOSING ASSERTION, and what makes the whole guard self-maintaining rather than three lists that
        // drift apart. A range-aware function is protected if EITHER the sweeping probe sees it OR a test
        // names it. This computes the complement — {Consumes entries blind to the whole sweep} minus {every
        // name asserted Consumes by either test} — and requires it to be EMPTY.
        //
        // WHAT THIS GUARANTEES, precisely, because the distinction matters and an earlier version of this
        // comment overstated it. It walks the Consumes half only, so it certifies nothing about an entry
        // flagged Elementwise; the two mistakes are protected by different tests:
        //
        //   * a FORGOTTEN flag is safe by construction — Consumes is the enum's zero value, so a new
        //     range-aware built-in registered with the default Entry<T> factory is never lifted. Nothing has
        //     to catch it. What this test adds is the guarantee that such an entry is still WATCHED: if the
        //     sweep cannot see it, it must be named above, and if it is neither, THIS test fails carrying its
        //     own name.
        //   * a WRONG flag — Elementwise on a range-aware built-in, the mistake that ships a silent wrong
        //     number — is caught by FunctionRegistryClassificationTests
        //     .TheElementwiseSet_IsExactlyTheCommittedRoster, which pins the Elementwise set BY NAME and so
        //     fails naming the newcomer, and (for anything the sweep can see) by
        //     EveryElementwiseEntry_IsBlindToTheRangeItIsHanded with a diagnostic. For the 21 entries the
        //     sweep is blind to, the roster is the ONLY defence, which is why it is pinned by name below
        //     rather than by a count.
        //
        // Verified by mutation: deleting the [Arguments("AND")] row above fails this test with
        // "Expected to be empty but collection contains items: [AND]".
        //
        // The named set is read by REFLECTION off the two tests' [Arguments] rows, so there is deliberately no
        // third copy of the list to fall out of sync — adding a row to either test is what registers a name
        // here. The two are TheShapeAndPositionAndCriteriaFamilies_StayConsumes above and
        // FunctionRegistryClassificationTests.TheHandAddedExclusions_StayConsumes, which owns the
        // two-population statistics, the paired-array sums and PERCENTILE.EXC/TRIMMEAN.
        var named = new HashSet<string>(
            [
                .. NamedConsumers(
                    typeof(ElementwiseLiftingTests),
                    nameof(TheShapeAndPositionAndCriteriaFamilies_StayConsumes)
                ),
                .. NamedConsumers(
                    typeof(Parsing.FunctionRegistryClassificationTests),
                    nameof(
                        Parsing
                            .FunctionRegistryClassificationTests
                            .TheHandAddedExclusions_StayConsumes
                    )
                ),
            ],
            StringComparer.Ordinal
        );

        var (workbook, sheet) = ThreeRectangles();
        var blind = new List<string>();
        var unnamed = new List<string>();

        foreach (var entry in FunctionRegistry.ByName.Values)
        {
            // A zero-argument entry (PI, RAND, …) has no argument to hand a rectangle to and can never lift,
            // which FunctionRegistryClassificationTests pins from the registry side.
            if (entry.Lifting is not ArrayLifting.Consumes || entry.MaxArgs is 0)
            {
                continue;
            }

            if (EveryPositionProbe(entry, workbook, sheet) is not null)
            {
                continue; // the probe sees it — a mis-flag would fail the oracle test with a diagnostic
            }

            blind.Add(entry.Name);

            if (!named.Contains(entry.Name))
            {
                unnamed.Add(entry.Name);
            }
        }

        await Assert.That(unnamed).IsEmpty();

        // Anti-vacuity, from both ends. The blind set is pinned BY NAME, not by a count: a count is
        // satisfiable by a compensating swap (one entry leaving the blind set as another joins it), which is
        // exactly the weakness TheElementwiseSet_IsExactlyTheCommittedRoster exists to close on the other
        // half. Pinning the names says out loud which 21 entries the sweep cannot see and therefore depend
        // ENTIRELY on being named above — and it fails, naming the newcomer, if a change to the probe or to
        // the registry adds one.
        await Assert
            .That(string.Join(", ", blind.Order(StringComparer.Ordinal)))
            .IsEqualTo(
                "AND, AREAS, FORECAST, FORECAST.LINEAR, FORMULATEXT, IF, INDIRECT, IRR, ISFORMULA, "
                    + "ISREF, LET, MIRR, OFFSET, OR, PERCENTILE.EXC, PROB, RANDBETWEEN, SHEET, TRIMMEAN, "
                    + "TYPE, XNPV"
            );

        // … and the loop really walked all 126 Consumes entries.
        await Assert
            .That(FunctionRegistry.ByName.Values.Count(e => e.Lifting is ArrayLifting.Consumes))
            .IsEqualTo(126);
    }

    // The [Arguments] rows of a one-string-parameter test, read back as the set of function names it asserts.
    private static IEnumerable<string> NamedConsumers(Type suite, string test) =>
        suite
            .GetMethod(test)!
            .GetCustomAttributes<ArgumentsAttribute>()
            .Select(arguments => (string)arguments.Values[0]!);
}
