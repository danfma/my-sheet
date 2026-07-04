using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase C of the mini-CSE plan: proves the volatile taint rises OUT of the element-wise
/// <see cref="Danfma.MySheet.Expressions.ArrayEvaluation"/> exactly as it does for a scalar formula.
///
/// <para>The mechanism (see <c>Workbook.GetCellValue</c>): the taint is a thread-local flag
/// (<c>_volatileTouched</c>), not a per-cell property. A volatile node (RAND/NOW/…) sets it via
/// <c>MarkVolatileTouched</c> during its synchronous <c>Evaluate</c>; the enclosing cell's
/// <c>GetCellValue</c> reads the flag after evaluating and, if set, records the cell key in the tainted set
/// so <c>Recalculate()</c> drops exactly those cells (and only those). The mini-CSE evaluator recurses
/// synchronously on the SAME thread inside that cell frame — a closed range's cells go through
/// <c>GetCellValue</c> (their taint OR's up the stack), and a broadcast scalar sub-expression (e.g. a bare
/// <c>RAND()</c> in a branch or on the other side of a comparison) is evaluated with the ordinary
/// <c>Expression.Evaluate</c>, which calls <c>MarkVolatileTouched</c> directly into the cell frame. Either
/// way the flag reaches the enclosing <c>GetCellValue</c>, so the array-consuming cell is tainted.</para>
///
/// <para>The observable, oracle-free proof of taint is behavioral and precise: after <c>Recalculate()</c> a
/// volatile array-formula cell REFRESHES (its cache entry was dropped ⇒ it was in the tainted set), while a
/// structurally identical NON-volatile array-formula cell stays put (the array path does not spuriously
/// taint). Within a single epoch both are stable (cached).</para>
/// </summary>
public class MiniCseVolatileTaintTests
{
    private static Workbook ShowHide(int seed, params string[] flags)
    {
        var workbook = new Workbook { RandomSeed = seed };
        var sheet = workbook.Sheets.Add("Sheet1");
        for (var i = 0; i < flags.Length; i++)
        {
            sheet[$"B{i + 2}"] = new StringValue(flags[i]);
        }

        return workbook;
    }

    private static double Cell(Workbook workbook, string id) =>
        workbook.GetCellValue("Sheet1", id).AsObject() is double d ? d : double.NaN;

    // --- A broadcast RAND() inside the array branch taints the array-consuming cell ---

    [Test]
    public async Task Sum_OfIfArray_WithBroadcastVolatileBranch_IsTaintedAndRefreshes()
    {
        // =SUM(IF(B2:B5="Show", RAND(), 0)) → RAND() is broadcast to the two "Show" rows, so the sum is
        // RAND()*2. RAND() never passes through a cell here (it is a broadcast scalar operand), yet its
        // MarkVolatileTouched lands in the enclosing cell frame → A1 is tainted.
        var workbook = ShowHide(2, "Hide", "Show", "Hide", "Show");
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse(
            "=SUM(IF(B2:B5=\"Show\", RAND(), 0))",
            workbook["Sheet1"]
        );

        var first = Cell(workbook, "A1");
        var firstAgain = Cell(workbook, "A1"); // same epoch → cached, stable
        await Assert.That(firstAgain).IsEqualTo(first);
        await Assert.That(first).IsGreaterThan(0d); // two "Show" rows × a positive draw

        workbook.Recalculate();
        var second = Cell(workbook, "A1");
        // Refreshed ⇒ A1 was in the tainted set that Recalculate dropped (a scalar =RAND()*2 behaves the same).
        await Assert.That(second).IsNotEqualTo(first);
    }

    // --- A condition-side volatile (the plan's literal repro) taints AND drives the numeric result ---

    [Test]
    public async Task Sum_OfIfArray_WithVolatileConditionOperand_RefreshesAcrossEpochs()
    {
        // =SUM(IF(B2:B5=IF(RAND()>0.5,"Show","Hide"),1,0)) — the RAND-driven inner IF picks the label to
        // count. With B = Show/Show/Show/Hide the count is 3 for "Show" and 1 for "Hide", so a fresh draw
        // each epoch flips the result: seeing >1 distinct value across epochs proves the volatile is
        // re-evaluated (cell tainted) rather than served stale from the cache.
        var workbook = ShowHide(4, "Show", "Show", "Show", "Hide");
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse(
            "=SUM(IF(B2:B5=IF(RAND()>0.5,\"Show\",\"Hide\"),1,0))",
            workbook["Sheet1"]
        );

        var seen = new HashSet<double>();
        for (var i = 0; i < 40; i++)
        {
            seen.Add(Cell(workbook, "A1"));
            workbook.Recalculate();
        }

        // Every value is one of the two possible counts …
        await Assert.That(seen.IsSubsetOf(new HashSet<double> { 1d, 3d })).IsTrue();
        // … and the volatile actually moved it across epochs (would be a single stale value if untainted).
        await Assert.That(seen.Count).IsGreaterThan(1);
    }

    // --- Control: a NON-volatile array formula is NOT tainted — Recalculate leaves it untouched ---

    [Test]
    public async Task Sum_OfNonVolatileIfArray_IsStableAcrossRecalculate()
    {
        // =SUM(IF(B2:B5="Show",1,0)) = 2 — array path, but no volatile. Recalculate must NOT refresh it
        // (proves the taint is precise: the mini-CSE does not spuriously mark every array-consuming cell).
        var workbook = ShowHide(2, "Hide", "Show", "Hide", "Show");
        workbook["Sheet1"]["A1"] = ExpressionParser.Parse(
            "=SUM(IF(B2:B5=\"Show\",1,0))",
            workbook["Sheet1"]
        );

        var first = Cell(workbook, "A1");
        await Assert.That(first).IsEqualTo(2d);

        workbook.Recalculate();
        var second = Cell(workbook, "A1");
        await Assert.That(second).IsEqualTo(first);
    }

    // --- A volatile cell READ THROUGH the array (a closed range holding a volatile) also taints ---

    [Test]
    public async Task Sum_OfIfArray_OverRangeWithVolatileCell_RefreshesAcrossRecalculate()
    {
        // B2:B5 carries a volatile cell (B3 = RAND()); the range comparison reads it via GetCellValue, so
        // its taint OR's up into A1. Here the array condition B2:B5>0.5 selects rows whose RAND draw exceeds
        // the threshold and sums those draws — a value that changes every epoch when A1 is correctly tainted.
        var workbook = new Workbook { RandomSeed = 8 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["B2"] = ExpressionParser.Parse("=RAND()", sheet);
        sheet["B3"] = ExpressionParser.Parse("=RAND()", sheet);
        sheet["B4"] = ExpressionParser.Parse("=RAND()", sheet);
        sheet["B5"] = ExpressionParser.Parse("=RAND()", sheet);
        sheet["A1"] = ExpressionParser.Parse("=SUM(IF(B2:B5>0.5,B2:B5,0))", sheet);

        var first = Cell(workbook, "A1");
        var firstAgain = Cell(workbook, "A1");
        await Assert.That(firstAgain).IsEqualTo(first); // stable within the epoch

        workbook.Recalculate();
        var second = Cell(workbook, "A1");
        await Assert.That(second).IsNotEqualTo(first); // volatile range cells tainted A1 → refreshed
    }
}
