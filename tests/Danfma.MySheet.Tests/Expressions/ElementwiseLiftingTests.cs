using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue; // TUnit also defines a StringValue

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 8 of <c>plans/structured-table-references-and-aggregate.md</c> — ELEMENTWISE LIFTING of unary
/// operators and pure-scalar built-ins inside the mini-CSE (<see cref="ArrayEvaluation"/>). Today the
/// evaluator recognizes only five array-producing shapes; a <c>UnaryOperation</c> or a scalar function over a
/// range falls to the opaque-scalar arm, is evaluated ONCE over a range (<c>#VALUE!</c>, since a range has no
/// scalar value) and broadcasts that error into every element — or, worse, has it silently CONSUMED by the
/// enclosing node (a comparison, IFERROR's error arm, COUNT's non-numeric skip), which is where the 0.0
/// answers pinned below come from.
///
/// <para>These are the phase's TDD pins, written BEFORE the engine change: every assertion whose comment says
/// "today" names the value measured on this tree at the time of writing, so the fix is a deliberate change of
/// answer rather than a silent one.</para>
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

        // (b) Over the plan's own E6:E8 = 1,2,3, whose golden it states as 3 (Excel: serial 1 is 1900-01-01,
        // so MONTH is 1,1,1 — Aspose CSE 3). This engine's serial 1 is 1899-12-31, so MONTH(1)=12,
        // MONTH(2)=1, MONTH(3)=1 and the lifted sum is 14. That epoch gap is a PRE-EXISTING divergence
        // recorded in the master plan's open decisions, not something this phase introduces or fixes: 14 is
        // pinned here so that closing the epoch gap updates this line deliberately. Today #VALUE!.
        await Assert.That(Num(OnNumeric("=SUM(MONTH(E6:E8))"))).IsEqualTo(14.0);
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
    public async Task ShapeMismatch_FillsEveryElementWithValueError()
    {
        // A 3x3 first argument against a 3x1 second: the mini-CSE's EXISTING dimension rule fills every
        // element with #VALUE!, so the fold reports #VALUE!. The lifted operand must obey that same rule
        // rather than inventing one of its own.
        //
        // KNOWN DIVERGENCE, pinned rather than asserted as Excel: real Excel BROADCASTS a 3x1 column across
        // a 3x3 (each row's value applied to every column). Measured against the P0 oracle, Aspose CSE:
        // SUM(LEFT(D7:F9,E6:E8)) = 0 (E6:E8 are blank here, so LEFT(...,0) = "" nine times), and isolated on
        // numbers SUM(A1:C3*E1:E3) = 108 and SUM(ROUND(A1:C3,E1:E3)) = 45. This engine answers #VALUE! for
        // all three TODAY, from the pre-existing BinaryOperand rule — measured: SUM(A1:C3*E1:E3) = #VALUE!
        // and COUNT(A1:C3*E1:E3) = 0, i.e. a per-ELEMENT fill, not a refusal of the whole expression. Excel
        // broadcasting is therefore a separate, pre-existing gap in the operand tree; this line keeps the
        // lift consistent with the engine it lands in, and closing the gap must update it deliberately.
        //
        // Note this assertion is already green today (#VALUE! from the opaque scalar rather than from the
        // fill) — it is a shape REGRESSION pin, and the ROUND(E6:E8,F6:F8) case above is what makes it mean
        // something.
        await Assert.That(OnTextual("=SUM(LEFT(D7:F9,E6:E8))")).IsEqualTo(ErrorValue.NotValue);
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
}
