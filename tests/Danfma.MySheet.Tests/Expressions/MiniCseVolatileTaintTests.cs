using Danfma.MySheet.Expressions;
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

    // --- Phase 8: a LIFTED node carries the taint exactly as the IF-array does ---

    [Test]
    public async Task Sum_OfLiftedFunction_WithBroadcastVolatileArgument_IsTaintedAndRefreshes()
    {
        // The lifted twin of Sum_OfIfArray_WithBroadcastVolatileBranch_… above. ROUND is an
        // ArrayLifting.Elementwise built-in, so =SUM(ROUND(A1:A3, RAND()*4)) builds a LiftedFunctionOperand
        // whose second argument is a broadcast ScalarOperand — RAND() is drawn ONCE at build time (the draw
        // count itself is pinned by ElementwiseLiftingMechanismTests) and its MarkVolatileTouched lands in
        // the enclosing cell frame, never through a cell. If the lift swallowed that, B1 would be untainted
        // and Recalculate would serve the same stale number forever.
        //
        // RAND()*4 spans num_digits 0..3, and the fixture's decimals differ at every one of those places, so
        // a fresh draw genuinely moves the sum.
        var workbook = new Workbook { RandomSeed = 3 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1.23456);
        sheet["A2"] = new NumberValue(2.34567);
        sheet["A3"] = new NumberValue(3.45678);
        sheet["B1"] = ExpressionParser.Parse("=SUM(ROUND(A1:A3,RAND()*4))", sheet);

        var first = Cell(workbook, "B1");
        await Assert.That(Cell(workbook, "B1")).IsEqualTo(first); // same epoch → cached, stable

        var seen = new HashSet<double>();
        for (var i = 0; i < 40; i++)
        {
            seen.Add(Cell(workbook, "B1"));
            workbook.Recalculate();
        }

        // Moved across epochs ⇒ B1 was in the tainted set Recalculate dropped.
        await Assert.That(seen.Count).IsGreaterThan(1);
    }

    [Test]
    public async Task Sum_OfNonVolatileLiftedFunction_IsStableAcrossRecalculate()
    {
        // The control, and the proof that the taint is PRECISE: the same lifted shape with a constant second
        // argument must NOT be marked. 1.23 + 2.35 + 3.46 = 7.04.
        var workbook = new Workbook { RandomSeed = 3 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1.23456);
        sheet["A2"] = new NumberValue(2.34567);
        sheet["A3"] = new NumberValue(3.45678);
        sheet["B1"] = ExpressionParser.Parse("=SUM(ROUND(A1:A3,2))", sheet);

        var first = Cell(workbook, "B1");
        await Assert.That(first).IsEqualTo(7.04);

        workbook.Recalculate();
        await Assert.That(Cell(workbook, "B1")).IsEqualTo(first);
    }

    [Test]
    public async Task Sum_OfLiftedUnary_OverAVolatileOperand_RefreshesAcrossEpochs()
    {
        // The unary half: -(A1:A3*RAND()) is a UnaryOperand over a BinaryOperand whose right side is the
        // broadcast volatile. Negating cannot lose the taint (it is a thread-local flag on the cell frame,
        // not a value property), and this pins that it does not.
        var workbook = new Workbook { RandomSeed = 5 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        sheet["B1"] = ExpressionParser.Parse("=SUM(-(A1:A3*RAND()))", sheet);

        var first = Cell(workbook, "B1");
        await Assert.That(Cell(workbook, "B1")).IsEqualTo(first);
        await Assert.That(first).IsLessThan(0d); // 6 * a positive draw, negated

        workbook.Recalculate();
        await Assert.That(Cell(workbook, "B1")).IsNotEqualTo(first);
    }

    // --- Phase 7: a PRODUCER's predicate vector is read at BUILD time, and the taint must still land ---

    [Test]
    public async Task Sum_OfFilter_WithAVolatileInclude_IsTaintedAndRefreshes()
    {
        // THE ONE PROPERTY IN THIS FILE THAT IS NOT ABOUT AN OPERAND'S VALUES. FILTER's include vector is
        // evaluated ONCE at BUILD time (the shape has to be constant before any element is read), which is
        // a different moment from every taint above: the volatile is drawn while the operand TREE is being
        // assembled, not while its elements are being streamed. If that moment sat outside the enclosing
        // cell's frame, MarkVolatileTouched would land nowhere, the cell would never enter the tainted set,
        // and Recalculate would serve the first draw's shape forever — the exact failure this pins against.
        //
        // FIXTURE, and why it is not the phase's 5/0/9. RAND() draws in (0,1), so A1:A3 = 5, 0, 9 makes the
        // include vector [TRUE, FALSE, TRUE] for EVERY possible draw and the sum a constant 14 — a value
        // that coincides with the right answer no matter whether the volatile is re-drawn, which pins
        // nothing at all. With A1:A3 = 0.25, 0.5, 0.75 the draw lands inside the data and decides the
        // SHAPE, so the answer is one of exactly four things:
        //     r <  0.25 → all three kept          1.5
        //     r <  0.5  → [0.5, 0.75]            1.25
        //     r <  0.75 → [0.75]                 0.75
        //     r >= 0.75 → nothing kept, no
        //                 if_empty              #CALC!
        // All four are binary-exact sums, and all four were observed across 40 epochs at seeds 1, 2 and 3
        // (measured 2026-09-10) — so "more than one distinct outcome" is a real signal here, not luck.
        var workbook = new Workbook { RandomSeed = 1 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(0.25);
        sheet["A2"] = new NumberValue(0.5);
        sheet["A3"] = new NumberValue(0.75);
        sheet["D1"] = ExpressionParser.Parse("=SUM(FILTER(A1:A3,A1:A3>RAND()))", sheet);

        // Stable inside one epoch (cached), exactly as a scalar volatile formula is.
        var first = workbook.GetCellValue("Sheet1", "D1").AsObject();
        await Assert.That(workbook.GetCellValue("Sheet1", "D1").AsObject()).IsEqualTo(first);

        var seen = new HashSet<object?>();
        for (var i = 0; i < 40; i++)
        {
            seen.Add(workbook.GetCellValue("Sheet1", "D1").AsObject());
            workbook.Recalculate();
        }

        // Moved across epochs ⇒ D1 was in the tainted set Recalculate dropped, so the build-time draw
        // reached the enclosing cell frame.
        await Assert.That(seen.Count).IsGreaterThan(1);

        // And every outcome is one of the four the fixture allows — the assertion that keeps "it moved"
        // from being satisfied by garbage.
        var allowed = new HashSet<object?> { 1.5, 1.25, 0.75, new ErrorValue("#CALC!") };
        await Assert.That(seen.IsSubsetOf(allowed)).IsTrue();
    }

    [Test]
    public async Task Sum_OfNonVolatileFilter_IsStableAcrossRecalculate()
    {
        // The control, and the proof the producer path does not taint indiscriminately: the same shape with
        // a CONSTANT threshold keeps [0.5, 0.75] = 1.25 across Recalculate. Without this, the test above
        // would also pass for an engine that marked every producer-bearing cell volatile.
        var workbook = new Workbook { RandomSeed = 1 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(0.25);
        sheet["A2"] = new NumberValue(0.5);
        sheet["A3"] = new NumberValue(0.75);
        sheet["D1"] = ExpressionParser.Parse("=SUM(FILTER(A1:A3,A1:A3>0.4))", sheet);

        var first = Cell(workbook, "D1");
        await Assert.That(first).IsEqualTo(1.25);

        workbook.Recalculate();
        await Assert.That(Cell(workbook, "D1")).IsEqualTo(first);
    }

    [Test]
    public async Task Sum_OfSequence_WithAVolatileArgument_RefreshesAcrossEpochs()
    {
        // SEQUENCE reads no cell at all, so this is the taint with NO range in the tree anywhere: the
        // volatile is a scalar argument of the producer itself, drawn once at build time, and the three
        // elements are start + 0*step. If the draw did not reach the cell frame the sum would be frozen.
        var workbook = new Workbook { RandomSeed = 11 };
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["D1"] = ExpressionParser.Parse("=SUM(SEQUENCE(3,1,RAND(),0))", sheet);

        var seen = new HashSet<double>();
        for (var i = 0; i < 20; i++)
        {
            seen.Add(Cell(workbook, "D1"));
            workbook.Recalculate();
        }

        await Assert.That(seen.Count).IsGreaterThan(1);
        // Every draw is in (0,1) and repeated three times with step 0, so every sum is in (0,3).
        await Assert.That(seen.All(value => value > 0d && value < 3d)).IsTrue();
    }

    [Test]
    public async Task AScalarConditionIf_WithAVolatileCondition_NeverCollapsesTheProducer()
    {
        // THE REGRESSION TEST FOR THE DEFECT THE FINAL REVIEW FOUND, and the reason it is here rather than
        // beside the other scalar-condition pins: it can only be seen by drawing the condition many times.
        // The first version of TryBuildScalarConditionIf DECLINED when the taken branch was a bare reference,
        // handing the choice back to the consumer's scalar path, which re-entered If.Evaluate and drew the
        // condition a SECOND time. When the two draws disagreed the fallback evaluated the OTHER branch as a
        // scalar, and a producer's scalar rule is its top-left — so SEQUENCE(3) became 1 and SUM answered 1.
        // Over 400 seeds that happened in exactly 100 of them. Nothing in the value pins could see it: every
        // single-seed run either takes the reference branch (#VALUE!) or agrees twice (6).
        //
        // So the assertion is over the SET of answers across seeds, and what makes it able to fail is the
        // third bucket: 1 must never appear. 200 seeds is enough — the defect hit one seed in four.
        var range = new HashSet<string>();
        var name = new HashSet<string>();

        for (var seed = 1; seed <= 200; seed++)
        {
            range.Add(Answer(seed, "=SUM(IF(RAND()<0.5,A1:A3,SEQUENCE(3)))"));
            name.Add(Answer(seed, "=SUM(IF(RAND()<0.5,Rng,SEQUENCE(3)))"));
        }

        // A1:A3 is 5, 0, 9. The reference branch is #VALUE! for a bare range and 14 for the name (the
        // scalar path's own answers, reproduced ONCE by WrapScalar); the producer branch is 6. A collapsed
        // producer would be "1" and a collapsed range would be "5".
        await Assert.That(range.OrderBy(x => x).ToArray()).IsEquivalentTo(["#VALUE!", "6"]);
        await Assert.That(name.OrderBy(x => x).ToArray()).IsEquivalentTo(["14", "6"]);

        // Both branches being producers never reached the declining path, which is exactly why an earlier
        // check of this method missed the defect. Kept as the contrast, not as the guard.
        var both = new HashSet<string>();
        for (var seed = 1; seed <= 200; seed++)
        {
            both.Add(Answer(seed, "=SUM(IF(RAND()<0.5,SEQUENCE(5),SEQUENCE(3)))"));
        }

        await Assert.That(both.OrderBy(x => x).ToArray()).IsEquivalentTo(["15", "6"]);

        static string Answer(int seed, string formula)
        {
            var workbook = new Workbook { RandomSeed = seed };
            var sheet = workbook.Sheets.Add("Sheet1");
            sheet["A1"] = new NumberValue(5);
            sheet["A2"] = new NumberValue(0);
            sheet["A3"] = new NumberValue(9);
            workbook.DefineName("Rng", "Sheet1!$A$1:$A$3");
            sheet["Z1"] = ExpressionParser.Parse(formula, sheet);

            var value = workbook.GetCellValue("Sheet1", "Z1").AsObject();

            return value is ErrorValue error ? error.ErrorCode : value?.ToString() ?? "null";
        }
    }

    [Test]
    public async Task AScalarConditionIf_OverAProducer_IsTaintedOnEveryBranchShape()
    {
        // The taint on the path Task 8 added, which had NO volatility coverage at all: every test above this
        // one reads IF(B2:B5…), an ARRAY condition. Measured over six Recalculate() passes on this build,
        // each of the three refreshes and the non-volatile control does not — the control is what stops this
        // test passing because everything refreshes.
        await Assert
            .That(DistinctOver("=SUM(IF(RAND()>0.5,SEQUENCE(3),SEQUENCE(5)))"))
            .IsGreaterThan(1);
        await Assert.That(DistinctOver("=SUM(IF(RAND()>0.5,SEQUENCE(3),0))")).IsGreaterThan(1);
        await Assert.That(DistinctOver("=SUM(IF(RAND()>0.5,1,2))")).IsGreaterThan(1);
        await Assert.That(DistinctOver("=SUM(IF(TRUE,SEQUENCE(3),0))")).IsEqualTo(1);

        static int DistinctOver(string formula)
        {
            var workbook = new Workbook { RandomSeed = 17 };
            var sheet = workbook.Sheets.Add("Sheet1");
            sheet["A1"] = new NumberValue(5);
            sheet["A2"] = new NumberValue(0);
            sheet["A3"] = new NumberValue(9);
            sheet["A1"] = ExpressionParser.Parse(formula, sheet);
            var seen = new HashSet<double>();

            for (var pass = 0; pass < 6; pass++)
            {
                seen.Add(Cell(workbook, "A1"));
                workbook.Recalculate();
            }

            return seen.Count;
        }
    }

    [Test]
    public async Task AnArrayConditionIf_WithAnOpenRangeBranch_IsRefusedAtTheProbe()
    {
        // The SECOND probe/build lockstep break the final review found, and the mirror image of the first:
        // ProbeBranches used to skip every bare reference node, so with an ARRAY condition it promised
        // "array" while the build — which really does construct BOTH branches for the zip — refused the open
        // range. The consumer then fell back and re-evaluated the condition, drawing a volatile TWICE where
        // the pre-Task-8 probe drew it once. The ANSWER never changed, which is why no value pin could see
        // it; the probe is now condition-aware instead.
        //
        // What makes this test able to fail is the pair: the open range must refuse and the CLOSED range
        // beside it must not. B1:B3 = 1, 2, 3 and A1:A3 = 0.2, 0.5, 0.8, so a closed branch sums some subset.
        await Assert.That(OnColumns("=SUM(IF(A1:A3>0,B:B,0))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnColumns("=SUM(IF(A1:A3>0,MyColumn,0))")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(OnColumns("=SUM(IF(A1:A3>0,B1:B3,0))")).IsEqualTo(6.0);

        // And the SCALAR-condition side must stay as the oracle has it, which is the reason the probe keys on
        // the condition rather than refusing everywhere: only one branch is ever built there, so a refusable
        // reference in the UNTAKEN branch costs nothing (measured 0 and 6 on the oracle), while a TAKEN open
        // range is the loud 1x1 #VALUE! WrapScalar gives it.
        await Assert.That(OnColumns("=SUM(IF(FALSE,MyColumn,0)*B1:B3)")).IsEqualTo(0.0);
        await Assert.That(OnColumns("=SUM(IF(TRUE,SEQUENCE(3),MyColumn))")).IsEqualTo(6.0);
        await Assert
            .That(OnColumns("=SUM(IF(TRUE,MyColumn,SEQUENCE(3)))"))
            .IsEqualTo(ErrorValue.NotValue);

        static object? OnColumns(string formula)
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.Add("Sheet1");
            sheet["A1"] = new NumberValue(0.2);
            sheet["A2"] = new NumberValue(0.5);
            sheet["A3"] = new NumberValue(0.8);
            sheet["B1"] = new NumberValue(1);
            sheet["B2"] = new NumberValue(2);
            sheet["B3"] = new NumberValue(3);
            workbook.DefineName("MyColumn", "Sheet1!$B:$B");

            return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
        }
    }
}
