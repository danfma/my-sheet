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
}
