using Danfma.MySheet.DirtyGraph;
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

    // --- Item 9: the classification guard (the regression defence for the next contributor) ---

    // A1:A3 = 1,2,3 and C4:C6 = 100,"x",blank — two rectangles that differ in POSITION and in CONTENTS, and
    // in the KIND of their contents. This is the executable form of the range-awareness oracle that produced
    // the 180/126 split, so the derivation re-runs on every build instead of trusting a frozen list.
    private static (Workbook Workbook, Sheet Sheet) TwoRectangles()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        sheet["C4"] = new NumberValue(100);
        sheet["C5"] = new StringValue("x");
        // C6 is deliberately BLANK — a third element kind the first rectangle does not have.

        return (workbook, sheet);
    }

    [Test]
    public async Task EveryElementwiseEntry_IsBlindToTheRangeItIsHanded()
    {
        // THE guard. For every entry the registry flags Elementwise, evaluate `F(A1:A3, 1, 1, …)` and
        // `F(C4:C6, 1, 1, …)` — MinArgs arguments, the rectangle in the first slot — on the ORDINARY scalar
        // path (Parse(...).Evaluate, which never enters the mini-CSE) and require the two answers to be
        // IDENTICAL. A pure-scalar body cannot tell the two rectangles apart: a range has no scalar value, so
        // it answers the same #VALUE! (or the same constant made from the other arguments) for both. A
        // RANGE-AWARE body reads the cells and therefore answers differently — 1,2,3 against 100,"x",blank —
        // which is exactly the mis-flag this test exists to catch, and the mis-flag that would otherwise ship
        // as a SILENT wrong number (the lift would hand the function one element of the rectangle it was
        // meant to consume whole; see the three measured silent cases in this file's class comment).
        //
        // Verified by mutation, twice. Flagging SUM `Elementwise` fails with
        // "SUM: A1:A3 → 6, C4:C6 → 100"; flagging COUNT — a range-aware function the explicit list below
        // does NOT name — fails with "COUNT: A1:A3 → 3, C4:C6 → 1". The second is the point: the oracle
        // catches a mis-flag nobody remembered to enumerate.
        //
        // Item 9(a) rides along: MinArgs arguments can only be built when the entry takes at least one, and a
        // zero-argument entry has nothing to lift (FunctionRegistryClassificationTests pins the same rule
        // from the registry side).
        var (workbook, sheet) = TwoRectangles();
        var lifted = new List<string>();
        var offenders = new List<string>();

        object? Answer(string body)
        {
            try
            {
                return ExpressionParser.ParseFormulaBody(body, sheet).Evaluate(workbook).AsObject();
            }
            catch (Exception exception)
            {
                // A throw is an answer too: the two rectangles must still reach the SAME one.
                return $"{exception.GetType().Name}: {exception.Message}";
            }
        }

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

            var filler = string.Concat(Enumerable.Repeat(",1", Math.Max(0, entry.MinArgs - 1)));
            var here = Answer($"{entry.Name}(A1:A3{filler})");
            var there = Answer($"{entry.Name}(C4:C6{filler})");

            if (!Equals(here, there))
            {
                offenders.Add($"{entry.Name}: A1:A3 → {here}, C4:C6 → {there}");
            }
        }

        await Assert.That(offenders).IsEmpty();
        // Anti-vacuity: the loop must actually have walked the whole lifted half of the registry.
        await Assert.That(lifted.Count).IsEqualTo(180);
    }

    [Test]
    public async Task EveryElementwiseEntryWithAnOptionalSlot_KeepsTheOmittedSlotUnderTheLift()
    {
        // Verifier correction M2, and the guard that the scalar-blindness test above cannot be: it compares
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
    public async Task TheShapeAndPositionAndCriteriaFamilies_StayConsumes(string name)
    {
        // The explicit half of the guard. ROWS/COLUMNS/AREAS/OFFSET/ISREF/TYPE are the family the executable
        // oracle above CANNOT see — they answer the SAME thing for two different rectangles (1x1 shapes, a
        // type code, a reference test), so scalar-blindness holds for them while they are still range-aware.
        // Naming them here is the only defence they have. IF and RANDBETWEEN are design exclusions (IF owns
        // a dedicated operand arm; lifting a volatile would draw once per element) and LET binds names to
        // whole sub-expressions.
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
}
