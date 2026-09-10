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
    public async Task CriteriaFamily_OverALiftedFunction_IsRef()
    {
        // The sixth consumer is still the ODD ONE OUT — CriteriaScan.Open is deliberately NOT
        // CriteriaScan.OpenArrayOrRange, so SUMIFS/COUNTIFS/… never STREAM a computed array — but Phase 11a
        // Rule B changed what they answer instead of streaming it: a lifted pure-scalar built-in in a range
        // slot is now REJECTED with #REF! (PositionalRange.RejectComputedArray) rather than collapsed to one
        // #VALUE! element and scanned. Widening the eligible set still does not make them read arrays.
        //
        // Aspose.Cells 26.6.0, re-measured 2026-09-10 on this fixture: #REF! array-entered (#VALUE! entered
        // plainly — a different mode, never compared against this one). The pin read ErrorValue.NotValue
        // (#VALUE!) up to 0b93d66, when the collapsed element WAS the answer's source; the mini-CSE
        // implements the array-entered rule, so #REF! is the value this engine owes.
        await Assert
            .That(OnTextual("=SUMIFS(LEN(A1:A3),A1:A3,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
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
    public async Task CriteriaFamily_OverABroadcastArray_IsRef()
    {
        // The sixth consumer stays the odd one out after Phase 10, and Phase 11a Rule B makes that explicit:
        // CriteriaScan.Open is deliberately NOT CriteriaScan.OpenArrayOrRange, so SUMIF/SUMIFS/COUNTIF never
        // stream a computed array — broadcast or not — and now REJECT it with #REF! rather than answering
        // from the collapsed scalar path (0, #VALUE!, 0 and 0, #VALUE!, 0 at 0b93d66, the divergence this
        // test used to pin deliberately).
        //
        // Aspose.Cells 26.6.0, measured 2026-09-10: all six lines below are #REF! array-entered (#VALUE!
        // entered plainly — a different mode, never compared against these). The composite half is what
        // Task 3's opening of the composite path made worth pinning: the criteria family must not start
        // seeing composites either, and rejecting them is how it does not.
        await Assert
            .That(OnBroadcastGrid("=SUMIF(A1:C3*H1:H2,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnBroadcastGrid("=SUMIFS(A1:C3,A1:C3*H1:H2,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnBroadcastGrid("=COUNTIF(A1:C3*H1:H2,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);

        // The same three over a COMPOSITE read at a foreign extent.
        await Assert
            .That(OnBroadcastGrid("=SUMIF((A1:A3*H1:H2)*E5:G5,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnBroadcastGrid("=SUMIFS(A1:C3,(A1:A3*H1:H2)*E5:G5,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnBroadcastGrid("=COUNTIF((A1:A3*H1:H2)*E5:G5,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
    }

    // ================================================================================================
    // Phase 7 — the PRODUCER direction. Everything above hands a consumer an array built from a RANGE
    // (an IF, an operator, a lifted call); everything below hands it an array a FUNCTION produced, which
    // is the direction this file had no coverage of at all. The four producers are FILTER, SORT, UNIQUE
    // and SEQUENCE, and they reach the consumers through the same ArrayEvaluation.TryStream gate — so
    // what these pins certify is that no consumer needed to learn anything about them.
    //
    // ORACLE. Aspose.Cells 26.6.0, re-measured 2026-09-10 for every value in this section, in BOTH entry
    // modes — plain (Cell.Formula) and array-entered (Cell.SetArrayFormula(f, 1, 1)). The two agree on
    // every row here EXCEPT the ones named in place; where they split, the ARRAY-ENTERED column is the
    // target, because the mini-CSE implements the array-entered rule everywhere. Modes are never mixed
    // inside one assertion, and no comment below compares one mode's number against the other's.
    //
    // THREE ROWS WHERE THE PIN IS NOT THE ORACLE'S NUMBER, each argued where it stands: the open-range
    // refusal (this phase's declared deviation), AVERAGE over UNIQUE (the oracle contradicts its own SUM
    // and COUNT, and the Microsoft AVERAGE page), and the scalar-condition IF (a MySheet defect, pinned
    // at today's wrong number so that fixing it turns those lines RED and names itself).
    // ================================================================================================

    // A1:A5 = 5, 0, 9, 0, 5. A1:A3 (5, 0, 9) is the phase fixture the composition rows were measured on;
    // A4/A5 repeat 0 and 5 so UNIQUE has something to remove (three distinct rows out of five cells) and
    // SORT/FILTER over the same five cells stay distinguishable from it. No producer's output equals its
    // source here and no two of the four produce the same array, so a pin cannot pass by streaming the
    // wrong one:
    //   FILTER(A1:A3,A1:A3>0) = [5, 9]        SORT(A1:A3) = [0, 5, 9]
    //   UNIQUE(A1:A5)         = [5, 0, 9]     SEQUENCE(5) = [1, 2, 3, 4, 5]
    private static object? OnProducerGrid(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);
        sheet["A4"] = new NumberValue(0);
        sheet["A5"] = new NumberValue(5);

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private const string Filtered = "FILTER(A1:A3,A1:A3>0)";
    private const string Filtered5 = "FILTER(A1:A5,A1:A5>0)";

    // --- Each of the four inside every numeric consumer this file covers ---

    [Test]
    public async Task Filter_StreamsIntoEveryNumericConsumer()
    {
        // FILTER(A1:A3,A1:A3>0) = [5, 9]: sum 14, two numbers, mean 7, and INDEX can name the second.
        // Oracle, both modes: 14, 2, 7, 5, 9, 5, 9, 7, 9.
        await Assert.That(Num(OnProducerGrid($"=SUM({Filtered})"))).IsEqualTo(14.0);
        await Assert.That(Num(OnProducerGrid($"=COUNT({Filtered})"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid($"=AVERAGE({Filtered})"))).IsEqualTo(7.0);
        await Assert.That(Num(OnProducerGrid($"=MIN({Filtered})"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid($"=MAX({Filtered})"))).IsEqualTo(9.0);
        await Assert.That(Num(OnProducerGrid($"=SMALL({Filtered},1)"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid($"=LARGE({Filtered},1)"))).IsEqualTo(9.0);
        await Assert.That(Num(OnProducerGrid($"=MEDIAN({Filtered})"))).IsEqualTo(7.0);
        await Assert.That(Num(OnProducerGrid($"=INDEX({Filtered},2)"))).IsEqualTo(9.0);
    }

    [Test]
    public async Task Sort_StreamsIntoEveryNumericConsumer()
    {
        // SORT(A1:A3) = [0, 5, 9] — a PERMUTATION, so SUM/COUNT/AVERAGE cannot tell it from the source
        // range and only the ORDER-SENSITIVE consumers prove the sort ran. INDEX(...,1) = 0 is the one
        // that does it: the source's first cell is 5. Oracle, both modes: 14, 3, 14/3, 0, 9, 5, 5, 5, 0.
        await Assert.That(Num(OnProducerGrid("=SUM(SORT(A1:A3))"))).IsEqualTo(14.0);
        await Assert.That(Num(OnProducerGrid("=COUNT(SORT(A1:A3))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=AVERAGE(SORT(A1:A3))"))).IsEqualTo(14.0 / 3.0);
        await Assert.That(Num(OnProducerGrid("=MIN(SORT(A1:A3))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnProducerGrid("=MAX(SORT(A1:A3))"))).IsEqualTo(9.0);
        await Assert.That(Num(OnProducerGrid("=SMALL(SORT(A1:A3),2)"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=LARGE(SORT(A1:A3),2)"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=MEDIAN(SORT(A1:A3))"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=INDEX(SORT(A1:A3),1)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Unique_StreamsIntoEveryNumericConsumer_AndTheOracleMisreadsOneOfThem()
    {
        // UNIQUE(A1:A5) = [5, 0, 9]: five cells in, three rows out, so SUM 14 (not 19) and COUNT 3
        // (not 5) are what prove the duplicates were dropped. Oracle, both modes: 14, 3, 0, 9, 5, 5, 5, 9.
        await Assert.That(Num(OnProducerGrid("=SUM(UNIQUE(A1:A5))"))).IsEqualTo(14.0);
        await Assert.That(Num(OnProducerGrid("=COUNT(UNIQUE(A1:A5))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=MIN(UNIQUE(A1:A5))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnProducerGrid("=MAX(UNIQUE(A1:A5))"))).IsEqualTo(9.0);
        await Assert.That(Num(OnProducerGrid("=SMALL(UNIQUE(A1:A5),2)"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=LARGE(UNIQUE(A1:A5),2)"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=MEDIAN(UNIQUE(A1:A5))"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=INDEX(UNIQUE(A1:A5),3)"))).IsEqualTo(9.0);

        // AVERAGE IS THE ONE ROW IN THIS SECTION WHERE THE ORACLE IS NOT FOLLOWED, and it is not a close
        // call — the oracle contradicts ITSELF inside one workbook and one entry mode. Measured 26.6.0,
        // 2026-09-10, both modes: AVERAGE(UNIQUE(A1:A5)) = 3, while on the SAME array SUM = 14, COUNT = 3,
        // COUNTA = 3, ROWS = 3, MEDIAN = 5, SUMPRODUCT = 14 and INDEX 1/2/3 = 5/0/9. A mean of 3 cannot
        // live with a sum of 14 over three numbers, and SUM(UNIQUE(..))/COUNT(UNIQUE(..)) written out in
        // the same sheet answers 4.666…, so the oracle disagrees with its own two halves.
        //
        // What the oracle actually does, isolated across six fixtures on 2026-09-10 (plain entry, one
        // formula per workbook): AVERAGE over a UNIQUE result drops every element up to and including the
        // last ZERO from its numerator while keeping the full count in its denominator —
        //   5,0,9 → 3      (= 9/3,  correct 14/3)        4,0,8 → 2.666… (= 8/3, correct 4)
        //   1,0,2 → 0.666… (= 2/3,  correct 1)           5,0   → 0      (= 0/2, correct 2.5)
        //   5,1,9 → 5      (= 15/3, CORRECT — no zero)   1..6  → 3.5    (CORRECT — no zero)
        // Remove the zero and the oracle is right, which is what identifies this as a defect in AVERAGE's
        // accumulation rather than a semantic of UNIQUE: no rule about distinct rows could depend on
        // whether one of them happens to be 0.
        //
        // So the page decides — P0's addendum says a measurement beats a page, but a measurement that
        // contradicts a measurement is not one. support.microsoft.com "AVERAGE function", fetched
        // 2026-09-10: the mean is "calculated by adding a group of numbers and then dividing by the count
        // of those numbers", and "cells with the value zero are included". [5, 0, 9] therefore averages
        // 14/3, which is what MySheet answers and what is pinned. Same clause the phase already used for
        // UNIQUE's exactly_once, where the oracle also contradicted its own row count.
        await Assert.That(Num(OnProducerGrid("=AVERAGE(UNIQUE(A1:A5))"))).IsEqualTo(14.0 / 3.0);

        // The control that keeps the paragraph above honest rather than a story: with the zero replaced by
        // a 1 the oracle agrees with MySheet at 5 (measured, both modes), so the rule is pinned on both
        // sides of its trigger.
        await Assert.That(Num(AverageOverUniqueWithoutAZero())).IsEqualTo(5.0);
    }

    // A1:A3 = 5, 1, 9 — the no-zero twin of the AVERAGE row above, where the oracle answers 5 as MySheet
    // does. A helper rather than a second shared fixture, so the trigger stays beside its pin.
    private static object? AverageOverUniqueWithoutAZero()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(1);
        sheet["A3"] = new NumberValue(9);

        return ExpressionParser
            .Parse("=AVERAGE(UNIQUE(A1:A3))", sheet)
            .Evaluate(workbook)
            .AsObject();
    }

    [Test]
    public async Task Sequence_StreamsIntoEveryNumericConsumer()
    {
        // SEQUENCE(5) = [1, 2, 3, 4, 5] and reads NO cell at all, so every number below comes from the
        // producer itself — the one consumer row in this section that cannot be satisfied by leaking a
        // range through. Oracle, both modes: 15, 5, 3, 1, 5, 2, 4, 3, 4.
        await Assert.That(Num(OnProducerGrid("=SUM(SEQUENCE(5))"))).IsEqualTo(15.0);
        await Assert.That(Num(OnProducerGrid("=COUNT(SEQUENCE(5))"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=AVERAGE(SEQUENCE(5))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=MIN(SEQUENCE(5))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnProducerGrid("=MAX(SEQUENCE(5))"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=SMALL(SEQUENCE(5),2)"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid("=LARGE(SEQUENCE(5),2)"))).IsEqualTo(4.0);
        await Assert.That(Num(OnProducerGrid("=MEDIAN(SEQUENCE(5))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=INDEX(SEQUENCE(5),4)"))).IsEqualTo(4.0);
    }

    // --- Composition: the design's main claim, that the recursive builder needed no new code ---

    [Test]
    public async Task AProducer_ComposesWithEveryOtherOperandKind()
    {
        // THE CHEAPEST PROOF THAT THE DESIGN'S CENTRAL CLAIM HOLDS: "a producer nested under a lifted
        // function, a unary, a binary or an IF, or under another producer, is reached by TryBuildOperand's
        // recursion" — no arm per combination. Ten rows, measured by this phase's first task and
        // RE-MEASURED here on 26.6.0, 2026-09-10 in both modes; all ten agree in plain and array-entered,
        // and all ten agree with this build, so nothing below re-pins a changed expectation.
        //
        //   lifted call over a producer        SUM(LEN(FILTER))                   2
        //   unary over a producer              SUM(-FILTER)                     -14
        //   producer in an IF condition        SUM(IF(FILTER>5,1,0))              1
        //   producer inside a producer         SUM(SORT(FILTER))                 14
        //   producer under an operator         SUM(FILTER*2)                     28
        //   shape gate over a nest             ROWS(UNIQUE(FILTER))               2
        //   INDEX over a nest with arguments   INDEX(SORT(FILTER,1,-1),1)         9
        //   producer as another's SOURCE       SUM(FILTER(SEQUENCE(5),SEQ>2))    12
        //   permutation then de-duplication    SUM(UNIQUE(SORT(A1:A3)))          14
        //   two producers of DIFFERENT sizes   SUM(SEQUENCE(3)*FILTER)         #N/A
        await Assert.That(Num(OnProducerGrid($"=SUM(LEN({Filtered}))"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(-{Filtered})"))).IsEqualTo(-14.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(IF({Filtered}>5,1,0))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(SORT({Filtered}))"))).IsEqualTo(14.0);
        await Assert.That(Num(OnProducerGrid($"=SUM({Filtered}*2)"))).IsEqualTo(28.0);
        await Assert.That(Num(OnProducerGrid($"=ROWS(UNIQUE({Filtered}))"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid($"=INDEX(SORT({Filtered},1,-1),1)"))).IsEqualTo(9.0);
        await Assert
            .That(Num(OnProducerGrid("=SUM(FILTER(SEQUENCE(5),SEQUENCE(5)>2))")))
            .IsEqualTo(12.0);
        await Assert.That(Num(OnProducerGrid("=SUM(UNIQUE(SORT(A1:A3)))"))).IsEqualTo(14.0);
        await Assert
            .That(OnProducerGrid($"=SUM(SEQUENCE(3)*{Filtered})"))
            .IsEqualTo(ErrorValue.NotAvailable);

        // Item 18's own three strings, over the five-cell fixture where the duplicates matter:
        // FILTER(A1:A5,A1:A5>0) = [5, 9, 5] → sorted 19, doubled 38; UNIQUE(A1:A5)'s first row is 5.
        // Oracle, both modes: 19, 38, 5.
        await Assert.That(Num(OnProducerGrid($"=SUM(SORT({Filtered5}))"))).IsEqualTo(19.0);
        await Assert.That(Num(OnProducerGrid($"=SUM({Filtered5}*2)"))).IsEqualTo(38.0);
        await Assert.That(Num(OnProducerGrid("=INDEX(UNIQUE(A1:A5),1)"))).IsEqualTo(5.0);
    }

    [Test]
    public async Task AProducer_InsideALiftedScalarFunction_IsLiftedPerElement()
    {
        // The LiftedFunctionOperand arm (Phase 8) over a producer child. LEN is the sharpest of these
        // because its answer has nothing to do with its input's magnitude: LEN over [0,5,9] is [1,1,1] = 3,
        // so a consumer that collapsed the producer to one element would answer 1. UNIQUE(A1:A5) is pinned
        // beside it at 3 — three rows out of five cells — while SUM over that same array is 14.
        // Oracle 26.6.0, 2026-09-10, both modes: 3, 6, 6, 8, 3, 2.
        await Assert.That(Num(OnProducerGrid("=SUM(LEN(SORT(A1:A3)))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=SUM(N(SEQUENCE(3)))"))).IsEqualTo(6.0);
        await Assert.That(Num(OnProducerGrid("=SUM(ABS(-SEQUENCE(3)))"))).IsEqualTo(6.0);

        // A lift with a SECOND, scalar argument broadcast across the producer: ROUND([5,9]/2, 0) is
        // [3, 5] (2.5 rounds away from zero, 4.5 to 5) = 8 — not 4, which is what one element would give.
        await Assert.That(Num(OnProducerGrid($"=SUM(ROUND({Filtered}/2,0))"))).IsEqualTo(8.0);

        // A lift over a de-duplicated array, and two lifts stacked over a producer.
        await Assert.That(Num(OnProducerGrid("=SUM(LEN(UNIQUE(A1:A5)))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(LEN(TRIM({Filtered})))"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task TwoProducers_UnderOneOperator_BroadcastByProjection()
    {
        // A producer on BOTH sides of a binary, where Phase 10's projection rule decides the answer rather
        // than either producer. Broadcasting.TryProject repeats an axis of extent 1 and answers #N/A for a
        // position no operand covers, so:
        //   3x1 against 2x1 → the third position is uncovered              #N/A
        //   2x1 against 2x1 → elementwise, 1*5 + 2*9                         23
        //   3x1 against 3x1 → 1*5 + 2*0 + 3*9                                32
        //   1x3 against 3x1 → a 3x3 outer product, (1+2+3)*(1+2+3)           36
        //   1x1 against 2x1 → the singleton repeats, 1*5 + 1*9               14
        //   2x1 against 1x2 → a 2x2, (5+9)*(1+2)                             42
        //   two DIFFERENT producers of the same shape, [0,5,9]·[5,0,9]       81
        //   5x1 against 3x1 → two positions uncovered                      #N/A
        // Oracle 26.6.0, 2026-09-10, every row identical in both modes. The two #N/A rows matter most: an
        // engine that silently recycled the shorter operand would answer a number.
        await Assert
            .That(OnProducerGrid($"=SUM(SEQUENCE(3)*{Filtered})"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Num(OnProducerGrid($"=SUM(SEQUENCE(2)*{Filtered})"))).IsEqualTo(23.0);
        await Assert.That(Num(OnProducerGrid("=SUM(SEQUENCE(3)*UNIQUE(A1:A5))"))).IsEqualTo(32.0);
        await Assert.That(Num(OnProducerGrid("=SUM(SEQUENCE(1,3)*SEQUENCE(3))"))).IsEqualTo(36.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(SEQUENCE(1)*{Filtered})"))).IsEqualTo(14.0);
        await Assert.That(Num(OnProducerGrid($"=SUM({Filtered}*SEQUENCE(1,2))"))).IsEqualTo(42.0);
        await Assert.That(Num(OnProducerGrid("=SUM(SORT(A1:A3)*UNIQUE(A1:A5))"))).IsEqualTo(81.0);
        await Assert
            .That(OnProducerGrid($"=SUM(SEQUENCE(5)*{Filtered5})"))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task AProducer_AsAnIfBranch_IsSelectedPerElement()
    {
        // A producer in an IF's BRANCH slot with an ARRAY condition — the shape the phase's own
        // re-verification names. IfOperand projects the condition and both branches together, so the
        // condition's extent chooses per element from the producer's:
        //   A1:A3>0 = [T,F,T] against SORT(A1:A3) = [0,5,9] → 0 + 0 + 9        9
        //   A1:A3>5 = [F,F,T], producer in the FALSE branch → 1 + 2 + 0        3
        //   producer in the CONDITION, [5,9]>5 = [F,T]      → 0 + 1            1
        //   condition 3x1 against a 2x1 producer branch     → uncovered     #N/A
        //   a producer on all three slots at once, [F,T,T]  → 0 + 2 + 3        5
        // Oracle 26.6.0, 2026-09-10, ARRAY-ENTERED column: 9, 3, 1, #N/A, 5. The first two split by entry
        // mode on the oracle (entered plainly they are #VALUE!, the pre-existing cell-boundary half
        // CellBoundaryIntersectionTests owns); this engine implements the array-entered rule, so that is
        // the column pinned, and the plain result is recorded only so nobody re-derives it and concludes a
        // row is missing.
        await Assert.That(Num(OnProducerGrid("=SUM(IF(A1:A3>0,SORT(A1:A3),0))"))).IsEqualTo(9.0);
        await Assert.That(Num(OnProducerGrid("=SUM(IF(A1:A3>5,0,SEQUENCE(3)))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(IF({Filtered}>5,1,0))"))).IsEqualTo(1.0);
        await Assert
            .That(OnProducerGrid($"=SUM(IF(A1:A3>0,{Filtered},0))"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Num(OnProducerGrid("=SUM(IF(SEQUENCE(3)>1,SEQUENCE(3),0))")))
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task AProducer_UnderAScalarConditionIf_CollapsesToItsTopLeft_ADivergence()
    {
        // A DIVERGENCE PINNED AT TODAY'S WRONG NUMBER, found while extending the IF row above. It is not
        // caused by this phase — but it is made SILENT by it, which is the part worth its own test.
        //
        // With a SCALAR condition, ArrayEvaluation's If arm does not build an array operand at all, so the
        // whole IF evaluates as an ordinary scalar expression. For a RANGE or a computed array that fails
        // LOUDLY, which is the pre-existing gap — measured on this build 2026-09-10, all three #VALUE!:
        //     SUM(IF(TRUE,A1:A3,0))   SUM(IF(TRUE,A1:A3*2,0))   SUM(IF(TRUE,LEN(A1:A3),0))
        // For a PRODUCER it does not fail: the producer's own scalar-context rule is "answer the top-left"
        // (this phase's item 3, FirstElement), so the sum is taken over ONE element and a wrong number is
        // served with no error at all:
        //     SUM(IF(TRUE,SEQUENCE(3),0))               1   (oracle 6,  both modes)
        //     SUM(IF(TRUE,SORT(A1:A3),0))               0   (oracle 14, both modes)
        //     SUM(IF(TRUE,FILTER(A1:A3,A1:A3>0),0))     5   (oracle 14, both modes)
        // The oracle is self-consistent here (6 = 1+2+3, 14 = 0+5+9) and the two entry modes agree, so
        // there is nothing to weigh: MySheet is wrong. The fix is an arm that lets a scalar-condition IF
        // still build its branch as an array — the same shape as correction M1's Let/Choose/unary-plus
        // arm, and owned by NO item in this phase.
        //
        // Pinned at the measured wrong value on purpose, following this file's own precedent for a
        // deviation: an expectation of 6 would sit RED with nobody assigned to it, while these three lines
        // turn red the moment someone fixes the arm, and the failure names the reason. If you are here
        // because they went red, that is the fix landing and the correct values are in this comment.
        await Assert.That(Num(OnProducerGrid("=SUM(IF(TRUE,SEQUENCE(3),0))"))).IsEqualTo(1.0);
        await Assert.That(Num(OnProducerGrid("=SUM(IF(TRUE,SORT(A1:A3),0))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnProducerGrid($"=SUM(IF(TRUE,{Filtered},0))"))).IsEqualTo(5.0);
    }

    [Test]
    public async Task AProducer_ThreeAndFourLevelsDeep_StillStreams()
    {
        // Depth, which is the property no pair of operands can demonstrate. FILTER(A1:A5,A1:A5>0) is
        // [5, 9, 5]; SORT makes it [5, 5, 9]; UNIQUE makes THAT [5, 9]. Every level changes the array, so
        // each number below is only reachable by running all of them:
        //   SUM(UNIQUE(SORT(FILTER)))               three producers        14  (19 without UNIQUE)
        //   INDEX(SORT(UNIQUE(FILTER),1,-1),1)      three, then INDEX       9
        //   SUM(LEN(UNIQUE(SORT(FILTER))))          three + a lift          2  (3 without UNIQUE)
        //   SUM(-SORT(UNIQUE(FILTER))*2)            three + unary + binary -28
        //   SUM(FILTER(SORT(A1:A5),SORT(A1:A5)>0))  a producer on BOTH of
        //                                           FILTER's own slots     19
        //   ROWS(UNIQUE(SORT(FILTER)))              the shape of the nest   2
        // Oracle 26.6.0, 2026-09-10, both modes: 14, 9, 2, -28, 19, 2.
        await Assert.That(Num(OnProducerGrid($"=SUM(UNIQUE(SORT({Filtered5})))"))).IsEqualTo(14.0);
        await Assert
            .That(Num(OnProducerGrid($"=INDEX(SORT(UNIQUE({Filtered5}),1,-1),1)")))
            .IsEqualTo(9.0);
        await Assert
            .That(Num(OnProducerGrid($"=SUM(LEN(UNIQUE(SORT({Filtered5}))))")))
            .IsEqualTo(2.0);
        await Assert
            .That(Num(OnProducerGrid($"=SUM(-SORT(UNIQUE({Filtered5}))*2)")))
            .IsEqualTo(-28.0);
        await Assert
            .That(Num(OnProducerGrid("=SUM(FILTER(SORT(A1:A5),SORT(A1:A5)>0))")))
            .IsEqualTo(19.0);
        await Assert.That(Num(OnProducerGrid($"=ROWS(UNIQUE(SORT({Filtered5})))"))).IsEqualTo(2.0);
    }

    // --- The gates this phase added (ROWS/COLUMNS) and the flattening family ---

    [Test]
    public async Task TheShapeAndFlatteningGates_SeeAProducersArray()
    {
        // ROWS/COLUMNS answer the producer's SHAPE without reading a value, and the flattening family
        // (COUNTA/CONCAT/TEXTJOIN) walks it element by element. SEQUENCE(2,3) is the row that proves
        // COLUMNS is not hard-wired to 1. Oracle 26.6.0, 2026-09-10, both modes for all of these.
        await Assert.That(Num(OnProducerGrid($"=ROWS({Filtered})"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid($"=COLUMNS({Filtered})"))).IsEqualTo(1.0);
        await Assert.That(Num(OnProducerGrid("=ROWS(SEQUENCE(2,3))"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid("=COLUMNS(SEQUENCE(2,3))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=ROWS(UNIQUE(A1:A5))"))).IsEqualTo(3.0);
        await Assert.That(Num(OnProducerGrid("=COLUMNS(SORT(A1:A3))"))).IsEqualTo(1.0);

        // COUNTA counts the produced ELEMENTS, so UNIQUE's 3 (from five cells) shows it is counting the
        // producer's output rather than its source.
        await Assert.That(Num(OnProducerGrid($"=COUNTA({Filtered})"))).IsEqualTo(2.0);
        await Assert.That(Num(OnProducerGrid("=COUNTA(SEQUENCE(5))"))).IsEqualTo(5.0);
        await Assert.That(Num(OnProducerGrid("=COUNTA(UNIQUE(A1:A5))"))).IsEqualTo(3.0);

        // CONCAT and TEXTJOIN expand every element, in order.
        await Assert.That(OnProducerGrid($"=CONCAT({Filtered})")).IsEqualTo("59");
        await Assert.That(OnProducerGrid($"=TEXTJOIN(\",\",TRUE,{Filtered})")).IsEqualTo("5,9");
        await Assert.That(OnProducerGrid("=TEXTJOIN(\"-\",TRUE,SEQUENCE(3))")).IsEqualTo("1-2-3");
    }

    [Test]
    public async Task TwoFunctionsDeliberatelyDoNotStreamAProducer()
    {
        // Symmetry with COUNTA/CONCAT/TEXTJOIN would be the WRONG expectation here, and this pin exists so
        // nobody "restores" it. Measured on the oracle 26.6.0, 2026-09-10, both entry modes:
        //   COUNTBLANK REJECTS a computed array with #REF! — the criteria family's answer, for every one
        //   of the four producers, rather than counting blanks in it;
        //   CONCATENATE takes the TOP-LEFT ("5", "1") where CONCAT expands ("59", "1-2-3" above).
        // Both match this build exactly, so both are pinned as measured behaviour, not as a deviation.
        await Assert
            .That(OnProducerGrid($"=COUNTBLANK({Filtered})"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnProducerGrid("=COUNTBLANK(SEQUENCE(3))"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(OnProducerGrid($"=CONCATENATE({Filtered})")).IsEqualTo("5");
        await Assert.That(OnProducerGrid("=CONCATENATE(SEQUENCE(3))")).IsEqualTo("1");
    }

    // --- The open-range refusal: this phase's declared deviation, and its silent half ---

    [Test]
    public async Task AProducerOverAnOpenRange_IsRefused_AndCountSwallowsTheRefusal()
    {
        // THE PHASE'S DECLARED DEVIATION, and the pin item 18 asks for by name. The mini-CSE refuses an
        // OpenRangeReference as an array operand (a cost guard, not a measurement), and a producer
        // propagates that refusal the way a binary or an IF does, so the consumer sees #VALUE!.
        //
        // Re-measured on the oracle 26.6.0, 2026-09-10, one formula per workbook with A1:A3 = 5, 0, 9 and
        // nothing else on the sheet: SUM(FILTER(A:A,A:A>0)) = 14 and ROWS(FILTER(A:A,A:A>0)) = 2 in BOTH
        // entry modes. The phase file records #VALUE! for the plain column of that first formula; that is
        // NOT reproducible on this version, so the deviation is a deviation in both modes and the older
        // note should not be repeated.
        await Assert.That(OnProducerGrid("=SUM(FILTER(A:A,A:A>0))")).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(OnProducerGrid("=ROWS(FILTER(A:A,A:A>0))"))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnProducerGrid("=SUM(SORT(A:A))")).IsEqualTo(ErrorValue.NotValue);

        // THE HALF THAT IS NOT LOUD, and the reason this pin covers more than SUM. COUNT and COUNTA
        // DISCARD errors instead of propagating them, so the refusal reaches them as a dropped element and
        // they answer a plausible number with no error at all: 0, 0, 0 and 1 on this build. Oracle,
        // ARRAY-ENTERED column: 3, 2, 3 and 3 (COUNT/COUNTA over SORT(A:A) split by entry mode on the
        // oracle; the array-entered column is the one quoted and the plain one is never compared against
        // it). Whoever narrows the open-range refusal must come back here — these four are the rows that
        // will not announce themselves.
        await Assert.That(Num(OnProducerGrid("=COUNT(UNIQUE(A:A))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnProducerGrid("=COUNT(FILTER(A:A,A:A>0))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnProducerGrid("=COUNT(SORT(A:A))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnProducerGrid("=COUNTA(SORT(A:A))"))).IsEqualTo(1.0);
    }
}
