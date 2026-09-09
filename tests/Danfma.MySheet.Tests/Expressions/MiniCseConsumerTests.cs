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
        // n beyond the 4-element vector → #REF! (Excel parity).
        await Assert.That(OnShowHide("=INDEX(ROW(B2:B5),5)")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(OnShowHide("=INDEX(ROW(B2:B5),0)")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Index_IntoWholeColumnRowNumbers_IsIdentity()
    {
        // ROW($A:$A) is the identity vector [1,2,3,…]; INDEX(…,n) = n, without materializing the column.
        await Assert.That(Num(OnShowHide("=INDEX(ROW($A:$A),4)"))).IsEqualTo(4.0);
        await Assert.That(Num(OnShowHide("=INDEX(ROW($A:$A),1)"))).IsEqualTo(1.0);
        // n < 1 is out of range → #REF!.
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
    public async Task Sum_OfRowOverRectangle_RepeatsEachRowNumberPerColumn()
    {
        // Regression pin for the pre-existing syntactic arm: ROW(A1:C3) is 3x3, each row number once per
        // COLUMN → 3*(1+2+3) = 18, not 1+2+3.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(A1:C3))"))).IsEqualTo(18.0);
    }

    [Test]
    public async Task Sum_OfColumnOverRange_IsTheSumOfTheColumnNumbers()
    {
        // COLUMN had no array operand at all before this fix: SUM saw the single leftmost column number 1.
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(A1:C1))"))).IsEqualTo(6.0);
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(A1:C3))"))).IsEqualTo(18.0);
        await Assert.That(Num(OnPositionGrid("=SMALL(COLUMN(A1:C1),2)"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task Sum_OfColumnOverName_RepeatsTheColumnPerRow()
    {
        // MyName is a single COLUMN of 3 rows: COLUMN(MyName) is [1,1,1] → 3 (and not 1, the scalar answer).
        await Assert.That(Num(OnPositionGrid("=SUM(COLUMN(MyName))"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task RowAndColumnOperands_AgreeOnRowMajorOrder()
    {
        // ROW(A1:B2) = [1,1,2,2] and COLUMN(A1:B2) = [1,2,1,2] in row-major order; zipped element-wise the
        // products are [1,2,2,4] → 9. Any disagreement on the traversal order (e.g. a column-major COLUMN)
        // still sums the same 4 values individually but pairs them differently: 1*1+1*2+2*1+2*2 = 9 only
        // holds for the row-major pairing, and this asserts the PAIRING, which the two SUM cases cannot.
        await Assert.That(Num(OnPositionGrid("=SUM(ROW(A1:B2)*COLUMN(A1:B2))"))).IsEqualTo(9.0);
    }

    [Test]
    public async Task SumProduct_OfRowAndColumn_ConsumesTheMiniCse()
    {
        // SUMPRODUCT opts into the mini-CSE too (PositionalRange gained an array backing): ROW(A1:B2) =
        // [1,1,2,2] and COLUMN(A1:B2) = [1,2,1,2] are zipped position by position, so this is the same 9
        // as the SUM form above — not the 1*1 of two collapsed scalars.
        await Assert
            .That(Num(OnPositionGrid("=SUMPRODUCT(ROW(A1:B2),COLUMN(A1:B2))")))
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
}
