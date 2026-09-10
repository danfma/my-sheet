using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase B of the mini-CSE plan: the production CONSUMERS (SUM-family via <c>NumericAggregation.Fold</c>,
/// SMALL/LARGE via <c>StatisticsMath.Collect</c>, and <c>INDEX</c>) opt into the element-wise
/// <see cref="ArrayEvaluation"/> for an array-eligible argument. Golden values are the K1 oracle repros
/// stated in <c>plans/mini-cse-array-arguments.md</c> (Excel implicit-intersection / CSE semantics).
/// </summary>
public class MiniCseConsumerTests
{
    // B2..B5 = Hide/Show/Hide/Show — the K1 repro fixture. Evaluates a formula against it.
    private static object? OnShowHide(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["B2"] = new StringValue("Hide");
        sheet["B3"] = new StringValue("Show");
        sheet["B4"] = new StringValue("Hide");
        sheet["B5"] = new StringValue("Show");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // D7="abc", D8="def", E9=" " (ONE space) — the Phase 8 Textual fixture (ElementwiseLiftingTests owns its
    // rationale). LEN over the 3x3 D7:F9 is [3,0,0 / 3,0,0 / 0,1,0]; after TRIM the single space collapses,
    // so the ninth-smallest is still 3 while the population loses its only 1.
    private static object? OnTextual(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["D7"] = new StringValue("abc");
        sheet["D8"] = new StringValue("def");
        sheet["E9"] = new StringValue(" ");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // A1:C3 grid plus the defined names the ROW/COLUMN array cases resolve THROUGH. The cell contents are
    // deliberate lies: no cell holds its own row or column number (so a position assertion can never be a
    // cell value leaking through), and A2 holds TEXT so that COUNT over MyName's CELLS is 2 while
    // COUNT(ROW(MyName)) — the element count of the row-number vector — is 3. MyColumn is the open-range
    // (cost-guard) twin of MyName, MyCell the single-cell one, and GhostName points at a sheet that does not
    // exist.
    private static object? OnPositionGrid(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["B1"] = new NumberValue(7);
        sheet["C1"] = new NumberValue(3);
        sheet["A2"] = new StringValue("x");
        sheet["B2"] = new NumberValue(8);
        sheet["C2"] = new NumberValue(6);
        sheet["A3"] = new NumberValue(9);
        sheet["B3"] = new NumberValue(4);
        sheet["C3"] = new NumberValue(11);

        workbook.DefineName("MyName", "Sheet1!$A$1:$A$3");
        workbook.DefineName("MyColumn", "Sheet1!$A:$A");
        workbook.DefineName("MyCell", "Sheet1!$A$2");
        workbook.DefineName("GhostName", "Ghost!$A$1:$A$3");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // --- SUM / COUNT / AVERAGE / MIN / MAX over an IF-array (NumericAggregation.Fold) ---

    [Test]
    public async Task Sum_OfIfArray_CountsMatchingBranch()
    {
        // =SUM(IF(B2:B5="Show",1,0)) → [0,1,0,1] → 2
        await Assert.That(Num(OnShowHide("=SUM(IF(B2:B5=\"Show\",1,0))"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Count_OfIfWithoutElse_IgnoresLogicalFalse()
    {
        // =COUNT(IF(B2:B5="Show",1)) → [FALSE,1,FALSE,1] → COUNT ignores the logicals → 2
        await Assert.That(Num(OnShowHide("=COUNT(IF(B2:B5=\"Show\",1))"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Average_OfIfArray_AveragesOnlyNumbers()
    {
        // =AVERAGE(IF(B2:B5="Show",ROW(B2:B5))) → [FALSE,3,FALSE,5] → mean{3,5} = 4
        await Assert
            .That(Num(OnShowHide("=AVERAGE(IF(B2:B5=\"Show\",ROW(B2:B5)))")))
            .IsEqualTo(4.0);
    }

    [Test]
    public async Task MinMax_OfIfArray_OverNumbersOnly()
    {
        // =MIN/MAX(IF(B2:B5="Show",ROW(B2:B5))) → {3,5}
        await Assert.That(Num(OnShowHide("=MIN(IF(B2:B5=\"Show\",ROW(B2:B5)))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnShowHide("=MAX(IF(B2:B5=\"Show\",ROW(B2:B5)))"))).IsEqualTo(5.0);
    }

    // --- SMALL over an IF-array (OrderSelection via StatisticsMath.Collect → Fold) ---

    [Test]
    public async Task Small_OfIfArray_FirstAndSecondVisibleRow()
    {
        // =SMALL(IF(B2:B5="Show",ROW(B2:B5)),k) → visible rows {3,5}
        await Assert
            .That(Num(OnShowHide("=SMALL(IF(B2:B5=\"Show\",ROW(B2:B5)),1)")))
            .IsEqualTo(3.0);
        await Assert
            .That(Num(OnShowHide("=SMALL(IF(B2:B5=\"Show\",ROW(B2:B5)),2)")))
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task Large_OfIfArray_KthLargestVisibleRow()
    {
        // =LARGE(IF(B2:B5="Show",ROW(B2:B5)),1) → largest of {3,5} = 5
        await Assert
            .That(Num(OnShowHide("=LARGE(IF(B2:B5=\"Show\",ROW(B2:B5)),1)")))
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task Small_OfIfArray_ErrorAfterKthElement_StillPropagates()
    {
        // Border (previously untested): a cell error sitting BEYOND the k-th selected value must still
        // propagate — the gather scans the whole array, and the first error (scan order) wins even when k is
        // already satisfiable. Here A2:A5 = [1, 2, 3, #DIV/0!]; SMALL(…,1) could "see" the min (1) at the
        // first position, yet the trailing #DIV/0! propagates (Excel parity). The heap-k streaming path must
        // preserve this exactly (scan every element; array error precedes any k/bounds outcome).
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A2"] = new NumberValue(1);
        sheet["A3"] = new NumberValue(2);
        sheet["A4"] = new NumberValue(3);
        sheet["A5"] = ExpressionParser.Parse("=1/0", sheet); // #DIV/0! at the LAST position
        sheet["B2"] = new StringValue("Show");
        sheet["B3"] = new StringValue("Show");
        sheet["B4"] = new StringValue("Show");
        sheet["B5"] = new StringValue("Show");

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        await Assert
            .That(Calc("=SMALL(IF(B2:B5=\"Show\",A2:A5),1)"))
            .IsEqualTo(ErrorValue.DivByZero);
        // LARGE takes the same path (min-heap); the trailing error propagates identically.
        await Assert
            .That(Calc("=LARGE(IF(B2:B5=\"Show\",A2:A5),1)"))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task KthValueStreaming_IgnoringErrors_SelectsOverThePostSkipPopulation()
    {
        // Phase 2 groundwork for AGGREGATE(15,6,…) — "SMALL, ignoring error values". AGGREGATE now DOES
        // pass ignoreErrors:true (Aggregate.ArrayForm, since 7d968a3), and this test stays as the DIRECT
        // contract of the flag on OrderSelection (the test project sees the assembly's internals): it pins
        // the two modes side by side on one stream, which no end-to-end AGGREGATE formula can do.
        // Fixture shape of Small_OfIfArray_ErrorAfterKthElement_StillPropagates:
        // A2:A5 = [1, 2, 3, #DIV/0!] behind an all-"Show" filter, so the stream is those four elements.
        //
        // ignoreErrors:true does ONE thing — the error element is not RECORDED (the scan still visits every
        // element) — and that alone yields AGGREGATE's contract, because `count` only ever counted the
        // NUMERIC elements: the population is {1,2,3}, so k=1 → 1, k=3 → 3, and k=4 is past the POST-skip
        // count → #NUM! (the four PRE-skip elements would have made 4 a legal k). ignoreErrors:false is the
        // unchanged SMALL behaviour: the trailing #DIV/0! wins over any k/bounds outcome.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A2"] = new NumberValue(1);
        sheet["A3"] = new NumberValue(2);
        sheet["A4"] = new NumberValue(3);
        sheet["A5"] = ExpressionParser.Parse("=1/0", sheet); // #DIV/0! at the LAST position
        sheet["B2"] = new StringValue("Show");
        sheet["B3"] = new StringValue("Show");
        sheet["B4"] = new StringValue("Show");
        sheet["B5"] = new StringValue("Show");

        var context = new EvaluationContext(workbook, sheet.Name);
        var array = ExpressionParser.Parse("=IF(B2:B5=\"Show\",A2:A5)", sheet);

        await Assert
            .That(ArrayEvaluation.TryEvaluateStream(array, context, out var stream))
            .IsTrue();

        var built = stream;

        object? Kth(double k, bool ignoreErrors) =>
            OrderSelection
                .KthValueStreaming(
                    built,
                    new NumberValue(k),
                    context,
                    largest: false,
                    ignoreErrors: ignoreErrors
                )
                .AsObject();

        await Assert.That(Num(Kth(1, ignoreErrors: true))).IsEqualTo(1.0);
        await Assert.That(Num(Kth(3, ignoreErrors: true))).IsEqualTo(3.0);
        await Assert.That(Kth(4, ignoreErrors: true)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Kth(1, ignoreErrors: false)).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Aggregate_OfIfArray_OptionSixSelectsOverThePostSkipPopulation()
    {
        // AGGREGATE as the fifth mini-CSE consumer, on the SAME fixture and the same array as
        // KthValueStreaming_IgnoringErrors_… above — but driven end to end through the engine node instead
        // of through OrderSelection directly. A2:A5 = [1, 2, 3, #DIV/0!] behind an all-"Show" filter.
        //
        // function_num 15 = SMALL. Option 4 ("ignore nothing") keeps SMALL's own propagation — the trailing
        // #DIV/0! wins over any k — while option 6 ("ignore error values") drops that element, leaving the
        // population {1,2,3}: k = 1 → 1, k = 3 → 3, and k = 4 is past the POST-skip count → #NUM! (the four
        // PRE-skip elements would have made 4 a legal k). This pair is the direct regression for the
        // ignoreErrors flag on the streaming heap.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A2"] = new NumberValue(1);
        sheet["A3"] = new NumberValue(2);
        sheet["A4"] = new NumberValue(3);
        sheet["A5"] = ExpressionParser.Parse("=1/0", sheet); // #DIV/0! at the LAST position
        sheet["B2"] = new StringValue("Show");
        sheet["B3"] = new StringValue("Show");
        sheet["B4"] = new StringValue("Show");
        sheet["B5"] = new StringValue("Show");

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        const string Array = "IF(B2:B5=\"Show\",A2:A5)";

        await Assert.That(Calc($"=AGGREGATE(15,4,{Array},1)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Num(Calc($"=AGGREGATE(15,6,{Array},1)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc($"=AGGREGATE(15,6,{Array},3)"))).IsEqualTo(3.0);
        await Assert.That(Calc($"=AGGREGATE(15,6,{Array},4)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Small_OfTheCorpusDivisionIdiom_PropagatesTheErrorElement()
    {
        // The corpus's "nth row that is neither blank nor zero" idiom, which divides BY the filter instead of
        // branching on it. On A1=5, A2=0, A3=9:
        //
        //   (A1:A3<>"")*(A1:A3<>0)                  → [1, 0, 1]   (the filter, A2 = 0 excluded)
        //   ROW(A1:A3) - ROW(INDEX(A1:A3,1,1)) + 1  → [1, 2, 3]   (positions RELATIVE to the block's top)
        //   the quotient                            → [1, #DIV/0!, 3]
        //
        // The subtrahend is the reference-returning-function case that stays SCALAR by design (documented
        // under "Not supported"), which is exactly what makes it usable here: ROW(INDEX(A1:A3,1,1)) is the
        // single number 1 broadcast across the vector, not a second [1,2,3].
        //
        // SMALL then reports the #DIV/0! for EVERY k, because the gather scans the whole array and the first
        // error wins (Small_OfIfArray_ErrorAfterKthElement_StillPropagates pins that rule) — which is
        // Excel-correct for SMALL. It is Phase 2's AGGREGATE(15,6,…) — function 15 = SMALL, option 6 = ignore
        // error values — that turns these two lines into 1 and 3; pinning the BEFORE here is what makes that
        // a deliberate change rather than a silent one.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        const string Quotient = "(ROW(A1:A3)-ROW(INDEX(A1:A3,1,1))+1)/((A1:A3<>\"\")*(A1:A3<>0))";

        await Assert.That(Calc($"=SMALL({Quotient},1)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc($"=SMALL({Quotient},2)")).IsEqualTo(ErrorValue.DivByZero);

        // The two assertions above would read the same if the vector were a single #DIV/0!, so the vector is
        // identified element-wise here. COUNT discards Fold's error channel (see Count.cs) and tallies the
        // two numbers that survive; the IF-shaped twin of the SAME filter never divides, so it yields the
        // worksheet rows 1 and 3 — the numerators that the quotient's 1 and 3 are the relative form of.
        await Assert.That(Num(Calc($"=COUNT({Quotient})"))).IsEqualTo(2.0);
        await Assert
            .That(Num(Calc("=SMALL(IF((A1:A3<>\"\")*(A1:A3<>0),ROW(A1:A3)),1)")))
            .IsEqualTo(1.0);
        await Assert
            .That(Num(Calc("=SMALL(IF((A1:A3<>\"\")*(A1:A3<>0),ROW(A1:A3)),2)")))
            .IsEqualTo(3.0);
    }

    [Test]
    public async Task Aggregate_OfTheCorpusDivisionIdiom_OptionSixSkipsTheErrorElement()
    {
        // The AFTER of the pin directly above: the SAME corpus idiom, on the SAME A1=5 / A2=0 / A3=9
        // fixture, is what AGGREGATE(15,6,…) turns from #DIV/0! into the answer the corpus wants. The
        // quotient is [1, #DIV/0!, 3]; function 15 = SMALL, option 6 = ignore error values, so the
        // population is {1,3} and k picks the 1st and the 2nd of it.
        //
        // k is written as the corpus writes it — ROWS($B$2:B2) / ROWS($B$2:B3) = 1 / 2, the expanding-range
        // counter a filled-down column uses — so the whole shape, not just the aggregate, is pinned.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        const string Quotient = "(ROW(A1:A3)-ROW(A1)+1)/((A1:A3<>\"\")*(A1:A3<>0))";

        await Assert.That(Num(Calc($"=AGGREGATE(15,6,{Quotient},ROWS($B$2:B2))"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc($"=AGGREGATE(15,6,{Quotient},ROWS($B$2:B3))"))).IsEqualTo(3.0);

        // Anti-vacuidade: sem a option 6 o mesmo shape continua sendo o #DIV/0! que o teste acima fixa.
        await Assert
            .That(Calc($"=AGGREGATE(15,4,{Quotient},ROWS($B$2:B2))"))
            .IsEqualTo(ErrorValue.DivByZero);
    }

    // --- INDEX over a materialized array first argument ---

    [Test]
    public async Task Index_IntoRowVector_ReturnsWorksheetRowNumber()
    {
        // =INDEX(ROW(B2:B5),1) → [2,3,4,5][1] = 2
        await Assert.That(Num(OnShowHide("=INDEX(ROW(B2:B5),1)"))).IsEqualTo(2.0);
        await Assert.That(Num(OnShowHide("=INDEX(ROW(B2:B5),3)"))).IsEqualTo(4.0);
    }

    [Test]
    public async Task Index_IntoRowVector_OutOfRange_IsRefError()
    {
        // n beyond the 4-element vector → #REF!, and that half IS Excel parity: Aspose.Cells 26.6.0 answers
        // #REF! for INDEX(ROW(B2:B5),5) too, plain and CSE-entered (measured 2026-09-09).
        //
        // n = 0 is a DIVERGENCE, not parity — the label this comment carried before was wrong. Aspose answers
        // 2: it intersects the WHOLE array and takes its first element, which for ROW(B2:B5) is B2's row
        // number (measured 2026-09-09, plain and CSE-entered alike). This engine rejects row_num < 1 outright
        // instead. Recorded for the planned Excel-compatibility sweep; the assertion below pins TODAY's
        // behaviour deliberately, so closing the gap has to be an explicit edit.
        await Assert.That(OnShowHide("=INDEX(ROW(B2:B5),5)")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(OnShowHide("=INDEX(ROW(B2:B5),0)")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Index_IntoWholeColumnRowNumbers_IsIdentity()
    {
        // ROW($A:$A) is the identity vector [1,2,3,…]; INDEX(…,n) = n, without materializing the column.
        await Assert.That(Num(OnShowHide("=INDEX(ROW($A:$A),4)"))).IsEqualTo(4.0);
        await Assert.That(Num(OnShowHide("=INDEX(ROW($A:$A),1)"))).IsEqualTo(1.0);
        // n < 1 → #REF! here, and — like the n = 0 case above — a DIVERGENCE rather than an out-of-range
        // rule: Aspose.Cells 26.6.0 answers 1, the identity vector's first element (measured 2026-09-09,
        // plain and CSE-entered). Same gap, same sweep; the assertion pins today's behaviour deliberately.
        await Assert.That(OnShowHide("=INDEX(ROW($A:$A),0)")).IsEqualTo(ErrorValue.Reference);
    }

    // --- The full BH25-like idiom: nested IF (with > and an IF-without-else), cross-sheet ---

    [Test]
    public async Task Bh25Like_NthVisibleRowNumber_CrossSheet()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        var calc = workbook.Sheets.Add("Calc");

        // Data!A2:A9 — visible rows (Show) are 2,4,5,7,9; keeping only ROW>2 leaves 4,5,7,9.
        string[] flags = ["Show", "Hide", "Show", "Show", "Hide", "Show", "Hide", "Show"];
        for (var i = 0; i < flags.Length; i++)
        {
            data[$"A{i + 2}"] = new StringValue(flags[i]);
        }

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, calc).Evaluate(workbook).AsObject();

        const string bh25 =
            "=INDEX(ROW($A:$A),SMALL(IF(Data!$A$2:$A$9=\"Show\","
            + "IF(ROW(Data!$A$2:$A$9)>2,ROW(Data!$A$2:$A$9))),{0}))";

        // The n-th matching row (Show AND row>2): {4,5,7,9}.
        await Assert.That(Num(Calc(string.Format(bh25, 1)))).IsEqualTo(4.0);
        await Assert.That(Num(Calc(string.Format(bh25, 2)))).IsEqualTo(5.0);
        await Assert.That(Num(Calc(string.Format(bh25, 3)))).IsEqualTo(7.0);
        await Assert.That(Num(Calc(string.Format(bh25, 4)))).IsEqualTo(9.0);
    }

    // --- Regressions: the dry-cell IF-array stays #VALUE!; pure ranges are unaffected ---

    [Test]
    public async Task DryCell_IfArray_StaysValueError()
    {
        // A cell whose whole formula is an array (no array-consuming function) keeps #VALUE! — Phase B
        // touches consumers only, never IF.Evaluate.
        await Assert.That(OnShowHide("=IF(B2:B5=\"Show\",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task PureNumericRange_AggregatesUnchanged()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        // SUM(range) stays on its Layer-2 memo path (arguments is [Reference]); SMALL(range,k) unchanged.
        await Assert.That(Num(Calc("=SUM(A1:A3)"))).IsEqualTo(6.0);
        await Assert.That(Num(Calc("=SMALL(A1:A3,1)"))).IsEqualTo(1.0);
        await Assert.That(Num(Calc("=INDEX(A1:A3,2)"))).IsEqualTo(2.0);
    }

    // --- ROW/COLUMN over a DEFINED NAME or any reference-producing node (the array half of FIX A) ---

    [Test]
    public async Task Sum_OfRowOverName_IsTheSumOfTheResolvedRowNumbers()
    {
        // MyName = Sheet1!$A$1:$A$3 → ROW(MyName) is the vector [1,2,3]; before this fix the name was an
        // opaque scalar and SUM saw the single top row number 1.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(MyName))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Count_OfRowOverName_CountsEveryElement()
    {
        // The sharpest sentinel: COUNT of the opaque scalar was 1 (a silently wrong answer, not an error).
        // 3 can only be the element count of the vector — MyName's own cells count 2 (A2 is text).
        await Assert.That(Num(OnPositionGrid("=COUNT(ROW(MyName))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnPositionGrid("=COUNT(MyName)"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Small_And_Index_OfRowOverName_SelectPositionally()
    {
        // The two other production consumers of the mini-CSE (OrderStatistics.KthValue, INDEX). k=1 alone
        // would be an assertion that cannot fail — the opaque scalar was the single value 1, whose 1st
        // smallest is also 1 — so k=3 carries the weight: over a one-element vector it was #NUM!.
        await Assert.That(Num(OnPositionGrid("=SMALL(ROW(MyName),1)"))).IsEqualTo(1.0);
        await Assert.That(Num(OnPositionGrid("=SMALL(ROW(MyName),3)"))).IsEqualTo(3.0);
        await Assert.That(Num(OnPositionGrid("=INDEX(ROW(MyName),2)"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Sum_OfRowOverRectangle_IsTheSumOfTheRowVector()
    {
        // ROW(A1:C3) is the 3x1 COLUMN [1,2,3] → 6. FLIPPED by Phase 10 (verifier correction B1): Phase 1
        // fabricated a 3x3 rectangle, each row number once per column, and pinned 18 here; Excel's shape is
        // the vector — Aspose.Cells 26.6.0, CSE column: 6, with COUNT(ROW(A1:C3)) = 3 pinned in
        // VectorBroadcastingTests. Plain entry does NOT agree, as an earlier version of this comment said:
        // typed, ROW of a rectangle is the single top row, so SUM and COUNT both answer 1 there (re-measured
        // 2026-09-10, both columns). The CSE column is the one this engine implements everywhere.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(A1:C3))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Sum_OfColumnOverRange_IsTheSumOfTheColumnNumbers()
    {
        // COLUMN had no array operand at all before this fix: SUM saw the single leftmost column number 1.
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(A1:C1))"))).IsEqualTo(6.0);
        await Assert.That(Num(OnPositionGrid("=SMALL(COLUMN(A1:C1),2)"))).IsEqualTo(2.0);

        // COLUMN(A1:C3) is the 1x3 ROW [1,2,3] → 6, the same as over A1:C1. FLIPPED by Phase 10 (verifier
        // correction B1) from the fabricated 3x3's 18 — Aspose.Cells 26.6.0, measured 2026-09-10, CSE column.
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(A1:C3))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task Sum_OfColumnOverName_IsTheOneColumnNumber()
    {
        // MyName is a single COLUMN of 3 rows: COLUMN(MyName) is the 1x1 vector [1] → 1. FLIPPED by Phase
        // 10 (verifier correction B1) from 3, the fabricated 3x1's [1,1,1] — Aspose.Cells 26.6.0, measured
        // 2026-09-10, CSE column. The row-vector twin, SUM(ROW(MyName)) = 6, is pinned above and did not
        // move; the pair is what shows the vector runs along the function's OWN axis.
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(MyName))"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task RowAndColumnOperands_AgreeOnRowMajorOrder()
    {
        // ROW(A1:B2) is the 2x1 column [1,2] and COLUMN(A1:B2) the 1x2 row [1,2]; the operator broadcasts
        // them into the 2x2 OUTER PRODUCT [1,2,2,4] → 9. Since Phase 10 this pins the outer product rather
        // than the pairing of two fabricated 2x2 rectangles (the same 9 by a different route), and the
        // asymmetric form that can tell the two apart lives in VectorBroadcastingTests:
        // SUM(ROW(A1:A3)*COLUMN(A1:C1)) = 36.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(A1:B2)*COLUMN(A1:B2))"))).IsEqualTo(9.0);
    }

    [Test]
    public async Task SumProduct_OfRowAndColumn_ConsumesTheMiniCse()
    {
        // SUMPRODUCT opts into the mini-CSE too (PositionalRange gained an array backing), but keeps its
        // own dimension rule: its ARGUMENTS must match exactly, and ROW(A1:B2) is a 2x1 against COLUMN's
        // 1x2 → #VALUE!. FLIPPED by Phase 10 (verifier correction B1) from 9, the zip of two fabricated 2x2
        // rectangles — Aspose.Cells 26.6.0, measured 2026-09-10, CSE column (plain entry agrees).
        await Assert
            .That(OnPositionGrid("=SUMPRODUCT(ROW(A1:B2),COLUMN(A1:B2))"))
            .IsEqualTo(ErrorValue.NotValue);

        // The working form: the broadcast happens INSIDE the one argument, which SUMPRODUCT then consumes
        // as the 2x2 outer product — 9, as the SUM form above.
        await Assert
            .That(Num(OnPositionGrid("=SUMPRODUCT(ROW(A1:B2)*COLUMN(A1:B2))")))
            .IsEqualTo(9.0);
    }

    [Test]
    public async Task Sum_OfRowOrColumnOverSingleCellName_StaysScalar()
    {
        // MyCell = Sheet1!$A$2 → a 1x1 reference is NOT an array (no shape to broadcast): the node evaluates
        // itself once and broadcasts, so ROW is its row 2 and COLUMN its column 1 (A2 holds text, which
        // never reaches the answer).
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(MyCell))"))).IsEqualTo(2.0);
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(MyCell))"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task Sum_OfRowOrColumnOverOpenRange_StaysScalar_CostGuard()
    {
        // The cost guard: an open column in an array position is REFUSED, so the consumer keeps its scalar
        // path and reads the DECLARED top row / leftmost column (1) instead of materializing 1,048,576 rows.
        // The literal form is refused syntactically; the NAME form is what pins the Refused arm of
        // ResolvePositionRange (MyColumn = Sheet1!$A:$A resolves to an OpenRangeReference).
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(A:A))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(A:A))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(MyColumn))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(MyColumn))"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task Sum_OfRowOverUnresolvableName_PropagatesTheNameError()
    {
        // An unresolved name is Scalar, not Refused: the consumer's scalar path evaluates ROW once, whose
        // own error recovery reports the argument's #NAME? rather than inventing a row number.
        await Assert.That(OnPositionGrid("=SUM(ROW(Nope))")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnPositionGrid("=SUM(COLUMN(Nope))")).IsEqualTo(ErrorValue.Name);

        // The case that actually distinguishes Scalar from Refused, and the one the oracle's comment cites:
        // nested in an operation with a real array, Scalar broadcasts the #NAME? into every element, while a
        // REFUSAL would collapse the whole operation to the scalar path, where a range in an addition is
        // #VALUE! — the argument's own error lost.
        await Assert.That(OnPositionGrid("=SUM(A1:A3+ROW(Nope))")).IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task Sum_OfRowOverNameOnMissingSheet_StaysRefError()
    {
        // A name standing for a range on a DELETED sheet is a structural #REF!, which the array path must
        // not turn into a plausible row number: resolving to a missing sheet degrades to Scalar so ROW's own
        // ReferenceGuard reports #REF! (the same re-check ReferencePosition does on the scalar path).
        await Assert.That(OnPositionGrid("=SUM(ROW(GhostName))")).IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnPositionGrid("=SUM(COLUMN(GhostName))"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap()
    {
        // The counterweight to the test above, and the one asymmetry between the syntactic fast-path arms
        // and the oracle: a LITERAL rectangle goes straight to a position operand, with no
        // ReferenceGuard.MissingSheet pass, while the oracle degrades a missing sheet to Scalar so ROW's own
        // guard reports #REF!. This is a KNOWN DIVERGENCE, not a rule — Excel answers #REF! for both — and it
        // is pinned only so that closing it (running the guard in the syntactic arms too) is a deliberate
        // edit that updates this line, instead of a silent change of answer. The same gap reaches the
        // anchored arm, whose comment in ArrayEvaluation.TryBuildPositionOperand cites this test.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(Ghost!A1:A3))"))).IsEqualTo(6.0);

        // The scalar path over the SAME reference is already right, which is what makes the divergence a gap
        // in the array path alone.
        await Assert.That(OnPositionGrid("=ROW(Ghost!A1:A3)")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Sum_OfRowOverDynamicRange_UsesTheResolvedSpan()
    {
        // The `or Reference` half of the arm, exercised by a ':' range with a reference-returning endpoint:
        // INDEX(A1:A3,1,1):A3 resolves to A1:A3, so ROW is [1,2,3] → 6. This is also the hook a structured
        // table-column node (Phase 3) will land on with no further edit.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(INDEX(A1:A3,1,1):A3))"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task ResolvableRowArgument_DrawsAnArgumentFunctionNoMoreOftenThanTheScalarPath()
    {
        // The probe and the build both RESOLVE the argument, so anything admitted to the resolvable arm is
        // resolved twice. That is why a reference-returning FUNCTION argument is deliberately left out: its
        // resolution would evaluate the function's own arguments, and TICK() would be drawn twice where the
        // scalar path draws it once. Measured here, and identical before this fix.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);

        var draws = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++draws);

        int DrawsOf(string formula)
        {
            draws = 0;
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook);
            return draws;
        }

        // ROW(INDEX(…)) — the excluded shape: one draw, on the opaque-scalar path.
        await Assert.That(DrawsOf("=SUM(ROW(INDEX(A1:A3,TICK(),1)))")).IsEqualTo(1);

        // A DynamicRange endpoint IS admitted, and costs no extra draw: the scalar path it replaces already
        // resolved twice (ReferenceGuard.MissingSheet, then ReferencePosition), so 2 is also the pre-fix count.
        await Assert.That(DrawsOf("=SUM(ROW(INDEX(A1:A3,TICK(),1):A3))")).IsEqualTo(2);
        await Assert.That(DrawsOf("=ROW(INDEX(A1:A3,TICK(),1):A3)")).IsEqualTo(2);
    }

    [Test]
    public async Task Sum_OfRowOverReferenceReturningFunction_StaysScalar()
    {
        // Deliberately NOT widened to a Function argument: probing INDEX/OFFSET/INDIRECT for its shape would
        // evaluate the function's own arguments once in the probe and again in the build (drawing a volatile
        // twice). ROW(INDEX(...)) therefore stays an opaque scalar here — the top row of the resolved
        // reference, 1 — exactly as before this fix; the scalar answer is correct, only the array shape is
        // deferred.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(INDEX(A1:A3,1,1)))"))).IsEqualTo(1.0);
    }

    // --- Phase 8: every consumer over a LIFTED argument (elementwise unary / pure-scalar built-in) ---

    [Test]
    public async Task Consumers_OverALiftedFunctionArgument_StreamItElementByElement()
    {
        // The mini-CSE has six consumers and Phase 8 changes what every one of them accepts: a pure-scalar
        // built-in over a range used to be an OPAQUE SCALAR (#VALUE!, broadcast), and is now an array. One
        // case per consumer family, all on the Textual fixture, all measured against the P0 oracle
        // (Aspose.Cells 26.6.0, 2026-09-09 — plain and CSE-entered agree on every line below).
        //
        // OrderSelection (SMALL/LARGE). The 1st smallest of [3,0,0,3,0,0,0,1,0] is 0 and the 9th is 3; the
        // PAIR is what makes this an array rather than a single value, since a one-element population makes
        // k=9 a #NUM!.
        await Assert.That(Num(OnTextual("=SMALL(LEN(D7:F9),1)"))).IsEqualTo(0.0);
        await Assert.That(Num(OnTextual("=SMALL(LEN(D7:F9),9)"))).IsEqualTo(3.0);

        // INDEX. Element 1 of the row-major vector is LEN("abc") = 3.
        await Assert.That(Num(OnTextual("=INDEX(LEN(D7:F9),1)"))).IsEqualTo(3.0);

        // AGGREGATE's array form, over two STACKED lifts (LEN over TRIM). Same k pair, same reasoning.
        await Assert.That(Num(OnTextual("=AGGREGATE(15,6,LEN(TRIM(D7:F9)),1)"))).IsEqualTo(0.0);
        await Assert.That(Num(OnTextual("=AGGREGATE(15,6,LEN(TRIM(D7:F9)),9)"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task CriteriaFamily_StillRefusesAComputedArray()
    {
        // The sixth consumer is the ODD ONE OUT and stays that way: CriteriaScan.Open (the criteria family's
        // entry point) is deliberately NOT CriteriaScan.OpenArrayOrRange, so SUMIFS/COUNTIFS/… never see a
        // computed array — only a real range. Widening the eligible set does not reach them.
        //
        // Aspose.Cells 26.6.0, measured 2026-09-09: #VALUE! entered plainly (and #REF! CSE-entered, which is
        // not the column this engine reproduces — the mini-CSE is never entered at the cell boundary).
        await Assert
            .That(OnTextual("=SUMIFS(LEN(A1:A3),A1:A3,\">0\")"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // --- Phase 10: every consumer over a BROADCAST argument, leaf pair and composite ---

    // Phase 10's fixture, the one VectorBroadcastingTests owns and documents: A1:C3 = 1..9 ROW-MAJOR,
    // E1:E3 = 1,2,3 (a 3x1 column), E5:G5 = 10,20,30 (a 1x3 row), H1:H2 = 1,2 (a 2x1 column). H1:H2 is the
    // only operand whose extent is neither 1 nor 3, so it is the one that leaves positions uncovered.
    private static object? OnBroadcastGrid(string formula)
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

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task OrderSelectionAndAggregate_OverABroadcastLeafPair()
    {
        // A1:C3 (3x3) against H1:H2 (2x1) is the phase's uncovered shape: [1,2,3 / 8,10,12 / #N/A x3].
        // Every value here is Aspose.Cells 26.6.0, measured 2026-09-10, CSE column. Green on arrival —
        // Task 3 completed the rule; these pins are the safety net that was missing, since the two SILENT
        // wrong numbers Task 3 found were exactly a shape no consumer pin covered.
        //
        // SMALL/LARGE propagate the error element (the gather scans the whole array and the first error
        // wins), so both are #N/A rather than the 1 and the 12 the covered population would give.
        await Assert
            .That(OnBroadcastGrid("=SMALL(A1:C3*H1:H2,1)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(OnBroadcastGrid("=LARGE(A1:C3*H1:H2,1)"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // AGGREGATE option 6 ("ignore error values") is the first consumer whose ANSWER changes from an
        // error to a NUMBER because of this phase: the three uncovered positions are skippable #N/A
        // elements where the pre-phase per-element #VALUE! fill covered all nine. Function 15 = SMALL,
        // 14 = LARGE; the post-skip population is the six covered numbers {1,2,3,8,10,12}, so k = 6 is its
        // largest (12), k = 1 of LARGE is the same 12, and k = 7 is past the end → #NUM!.
        await Assert.That(Num(OnBroadcastGrid("=AGGREGATE(15,6,A1:C3*H1:H2,6)"))).IsEqualTo(12.0);
        await Assert.That(Num(OnBroadcastGrid("=AGGREGATE(14,6,A1:C3*H1:H2,1)"))).IsEqualTo(12.0);
        await Assert
            .That(OnBroadcastGrid("=AGGREGATE(15,6,A1:C3*H1:H2,7)"))
            .IsEqualTo(ErrorValue.Number);

        // Function 9 = SUM takes the REFERENCE slot (Phase 2's rule), which a computed array can never
        // fill: #VALUE!, and this phase does not change that.
        await Assert
            .That(OnBroadcastGrid("=AGGREGATE(9,6,A1:C3*H1:H2)"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task EveryConsumer_OverACompositeReadAtAForeignExtent()
    {
        // The path Tasks 2 and 3 opened, and the one whose absence let two silent wrong numbers survive:
        // the consumer's argument is a COMPOSITE whose own extent is not the extent it is read at.
        //
        // (A1:A3*H1:H2) is a 3x1 composite [1, 8, #N/A]; times the 1x3 row E5:G5 it is read at 3x3 —
        // 10,20,30 / 80,160,240 / #N/A x3. (E1:E3+1) is a 3x1 composite [2,3,4] read at the same 3x3 with
        // NOTHING uncovered — 20,40,60 / 30,60,90 / 40,80,120 — so each consumer is pinned on both sides
        // of the rule and a pin cannot pass by answering "error" everywhere.
        //
        // Aspose.Cells 26.6.0, measured 2026-09-10, CSE column, every line. All green on arrival.
        const string Uncovered = "(A1:A3*H1:H2)*E5:G5";
        const string Covered = "(E1:E3+1)*E5:G5";

        // The SUM/COUNT frame the rest read against: six covered elements of nine.
        await Assert.That(OnBroadcastGrid($"=SUM({Uncovered})")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnBroadcastGrid($"=COUNT({Uncovered})"))).IsEqualTo(6.0);
        await Assert.That(Num(OnBroadcastGrid($"=SUM({Covered})"))).IsEqualTo(540.0);
        await Assert.That(Num(OnBroadcastGrid($"=COUNT({Covered})"))).IsEqualTo(9.0);

        // SMALL / LARGE.
        await Assert
            .That(OnBroadcastGrid($"=SMALL({Uncovered},1)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(OnBroadcastGrid($"=LARGE({Uncovered},1)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnBroadcastGrid($"=SMALL({Covered},1)"))).IsEqualTo(20.0);
        await Assert.That(Num(OnBroadcastGrid($"=LARGE({Covered},1)"))).IsEqualTo(120.0);

        // k = 9 over the covered composite is the assertion that cannot be satisfied by a one-element
        // population: the composite really is read nine times, not once.
        await Assert.That(Num(OnBroadcastGrid($"=SMALL({Covered},9)"))).IsEqualTo(120.0);

        // MAX / MIN — the NumericAggregation.Fold family, same stream as SUM.
        await Assert.That(OnBroadcastGrid($"=MAX({Uncovered})")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(OnBroadcastGrid($"=MIN({Uncovered})")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnBroadcastGrid($"=MAX({Covered})"))).IsEqualTo(120.0);
        await Assert.That(Num(OnBroadcastGrid($"=MIN({Covered})"))).IsEqualTo(20.0);

        // AGGREGATE. Option 6 skips the three uncovered elements, so the post-skip population is the six
        // covered numbers {10,20,30,80,160,240}: k = 1 → 10, k = 6 → 240, LARGE's k = 1 → 240, k = 7 past
        // the end → #NUM!. Function 9 keeps its reference-slot #VALUE! even where nothing is uncovered.
        await Assert.That(Num(OnBroadcastGrid($"=AGGREGATE(15,6,{Uncovered},1)"))).IsEqualTo(10.0);
        await Assert.That(Num(OnBroadcastGrid($"=AGGREGATE(15,6,{Uncovered},6)"))).IsEqualTo(240.0);
        await Assert.That(Num(OnBroadcastGrid($"=AGGREGATE(14,6,{Uncovered},1)"))).IsEqualTo(240.0);
        await Assert
            .That(OnBroadcastGrid($"=AGGREGATE(15,6,{Uncovered},7)"))
            .IsEqualTo(ErrorValue.Number);
        await Assert.That(Num(OnBroadcastGrid($"=AGGREGATE(15,6,{Covered},1)"))).IsEqualTo(20.0);
        await Assert
            .That(OnBroadcastGrid($"=AGGREGATE(9,6,{Covered})"))
            .IsEqualTo(ErrorValue.NotValue);

        // INDEX, which is the consumer that can name WHICH positions are uncovered.
        await Assert.That(Num(OnBroadcastGrid($"=INDEX({Uncovered},2,3)"))).IsEqualTo(240.0);
        await Assert
            .That(OnBroadcastGrid($"=INDEX({Uncovered},3,1)"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // SUMPRODUCT's composite cases live in MathAggregateTests, beside its own dimension rule.
    }

    [Test]
    public async Task CriteriaFamily_OverABroadcastArray_StillRefusesIt()
    {
        // The sixth consumer stays the odd one out after Phase 10 exactly as it was before it:
        // CriteriaScan.Open is deliberately NOT CriteriaScan.OpenArrayOrRange, so SUMIF/SUMIFS/COUNTIF
        // never see a computed array — broadcast or not — and answer from the scalar path instead.
        //
        // This is a DIVERGENCE, pinned deliberately, not a rule. Aspose.Cells 26.6.0, measured
        // 2026-09-10: all six lines below are #REF! CSE-entered (#VALUE! entered plainly). Closing it is
        // Phase 11's, and it has to be an edit to these lines rather than a silent change of answer.
        // Green on arrival, and the composite half is what Task 3's opening of the composite path makes
        // worth pinning: the criteria family must not start seeing composites either.
        await Assert.That(Num(OnBroadcastGrid("=SUMIF(A1:C3*H1:H2,\">0\")"))).IsEqualTo(0.0);
        await Assert
            .That(OnBroadcastGrid("=SUMIFS(A1:C3,A1:C3*H1:H2,\">0\")"))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Num(OnBroadcastGrid("=COUNTIF(A1:C3*H1:H2,\">0\")"))).IsEqualTo(0.0);

        // The same three over a COMPOSITE read at a foreign extent.
        await Assert
            .That(Num(OnBroadcastGrid("=SUMIF((A1:A3*H1:H2)*E5:G5,\">0\")")))
            .IsEqualTo(0.0);
        await Assert
            .That(OnBroadcastGrid("=SUMIFS(A1:C3,(A1:A3*H1:H2)*E5:G5,\">0\")"))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Num(OnBroadcastGrid("=COUNTIF((A1:A3*H1:H2)*E5:G5,\">0\")")))
            .IsEqualTo(0.0);
    }
}
