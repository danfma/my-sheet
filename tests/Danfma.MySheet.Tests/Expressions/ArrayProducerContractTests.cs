using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Text;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;
using Index = Danfma.MySheet.Expressions.Lookup.Index;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 7 (dynamic arrays), the PRODUCER CONTRACT: <see cref="IArrayProducer"/>, its two dispatch arms in
/// <see cref="ArrayEvaluation"/>, <see cref="ArrayEvaluation.FirstElement"/>, and the shared operands in
/// <c>ArrayShaping.cs</c> (<see cref="SingletonArrayOperand"/>, <see cref="AxisSelectionOperand"/>) with
/// their hard invariant. None of the four functions exists yet — these tests drive the contract through
/// TEST-ONLY producers, so that FILTER, SORT, UNIQUE and SEQUENCE plug into a channel that is already
/// proved to compose with the mini-CSE (recursion, broadcasting, lifting, IF, nesting, the cell boundary).
/// </summary>
/// <remarks>
/// <para>Oracle. Every golden that names a formula was MEASURED on Aspose.Cells 26.6.0 (2026-09-10) with
/// the four real producers over the same fixture — <c>A1:A3</c> = 5, 0, 9; <c>B1:B3</c> = 1, 2, 3;
/// <c>H1:H2</c> = 1, 2 — in BOTH entry modes; the test-only producer reproduces the SHAPE and VALUES the
/// real one yields, so the number is a pin of the contract, not of the function. Where the two modes
/// differ, the CSE column is the target (the mini-CSE implements the array-entered rule) and the plain
/// value is named; "both modes" means they agreed.</para>
/// <para>An assertion with no formula named is an engine-mechanism pin (which element a selection maps
/// where), not an oracle claim.</para>
/// </remarks>
public class ArrayProducerContractTests
{
    // ------------------------------------------------------------------ fixture and helpers

    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Number(5);
        sheet["A2"] = Number(0);
        sheet["A3"] = Number(9);
        sheet["B1"] = Number(1);
        sheet["B2"] = Number(2);
        sheet["B3"] = Number(3);
        sheet["H1"] = Number(1);
        sheet["H2"] = Number(2);

        return (workbook, sheet);
    }

    private static object? Eval(Expression node, Workbook workbook) =>
        node.Evaluate(workbook).AsObject();

    private static double Num(object? value) => value is double d ? d : double.NaN;

    private static double NumberAt(ArrayEvaluationResult result, int index)
    {
        result.Values[index].TryGetNumber(out var value);
        return value;
    }

    private static BinaryOperation Multiply(Expression left, Expression right) =>
        new(BinaryOperator.Multiply, left, right);

    // ------------------------------------------------------------------ test-only producers

    // A row-major literal array that projects through Broadcasting.TryProject exactly as every array
    // operand does — the shape of a SequenceOperand, with the values handed in. It has NO shape guard on
    // purpose: the zero-extent demonstration below needs to build the operand the invariant forbids.
    private sealed class LiteralArrayOperand(int rows, int columns, params double[] values)
        : ArrayOperand
    {
        public override bool IsArray => true;
        public override int Rows => rows;
        public override int Columns => columns;

        public override ComputedValue At(int index, int rows, int columns) =>
            Broadcasting.TryProject(index, rows, columns, Rows, Columns, out var own)
                ? ComputedValue.Number(values[own])
                : ComputedValue.Error(Error.NA);
    }

    // The smallest producer: hands back a prepared operand. `Eligible` false models a producer whose
    // ProbeArray propagates a refused child — and, in lockstep, whose build refuses too: the two members
    // are twins, and a stub whose probe refused while its build succeeded answered 1 for SUM(refused)
    // instead of #VALUE! when this file was first run (FirstElement trusts the BUILD).
    private sealed record StubProducer(
        Func<EvaluationContext, ArrayOperand> Build,
        bool Eligible = true
    ) : Expression, IArrayProducer
    {
        public override ComputedValue Evaluate(EvaluationContext context) =>
            ArrayEvaluation.FirstElement(this, context);

        (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
            Eligible ? (true, true) : (false, false);

        bool IArrayProducer.TryBuildArrayOperand(
            EvaluationContext context,
            out ArrayOperand operand
        )
        {
            if (!Eligible)
            {
                operand = null!;
                return false;
            }

            operand = Build(context);
            return true;
        }
    }

    private static StubProducer Literal(int rows, int columns, params double[] values) =>
        new(_ => new LiteralArrayOperand(rows, columns, values));

    // A producer over a CHILD expression — a fixed selection along one axis of whatever the child builds
    // to. This is the whole shape of FILTER/SORT/UNIQUE minus the rule that computes the selection: it
    // recurses through the widened ArrayEvaluation.Probe/TryBuildOperand, propagates a refused child as
    // its own refusal (the Binary/IF side, not the lift's opaque scalar), and wraps a scalar child in a
    // SingletonArrayOperand (blocker B1) so a 1x1 source is still an ARRAY result.
    private sealed record Select(Expression Source, ArrayAxis Axis, int[] Selection)
        : Expression,
            IArrayProducer
    {
        public override ComputedValue Evaluate(EvaluationContext context) =>
            ArrayEvaluation.FirstElement(this, context);

        (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
            ArrayEvaluation.Probe(Source, context).Succeeds ? (true, true) : (false, false);

        bool IArrayProducer.TryBuildArrayOperand(
            EvaluationContext context,
            out ArrayOperand operand
        )
        {
            if (!ArrayEvaluation.TryBuildOperand(Source, context, out var source))
            {
                operand = null!;
                return false;
            }

            operand = new AxisSelectionOperand(
                source.IsArray ? source : new SingletonArrayOperand(source.Scalar),
                Axis,
                Selection
            );
            return true;
        }
    }

    // ------------------------------------------------------------------ the two dispatch arms

    [Test]
    public async Task AProducer_IsArrayEligible_AndStreamsItsElements_ThroughBothArms()
    {
        var (workbook, _) = Grid();
        var context = new EvaluationContext(workbook);

        // The 3x1 array FILTER(A1:A3,TRUE) yields over this fixture: SUM 14, COUNT 3 (both modes,
        // DynamicArrayTests). Probe's arm says "array", the build's arm streams it — in lockstep.
        var producer = Literal(3, 1, 5, 0, 9);

        await Assert.That(ArrayEvaluation.IsArrayEligible(producer, context)).IsTrue();
        await Assert.That(ArrayEvaluation.TryEvaluate(producer, context, out var result)).IsTrue();
        await Assert.That(result.Rows).IsEqualTo(3);
        await Assert.That(result.Columns).IsEqualTo(1);
        await Assert.That(NumberAt(result, 0)).IsEqualTo(5.0);
        await Assert.That(NumberAt(result, 1)).IsEqualTo(0.0);
        await Assert.That(NumberAt(result, 2)).IsEqualTo(9.0);

        await Assert.That(Num(Eval(new Sum([producer]), workbook))).IsEqualTo(14.0);
        await Assert.That(Num(Eval(new Count([producer]), workbook))).IsEqualTo(3.0);
        await Assert.That(Num(Eval(new Index([producer, Number(3)]), workbook))).IsEqualTo(9.0);
    }

    [Test]
    public async Task ARefusedChild_PropagatesTheRefusal_AndTheScalarPathAnswersValueError()
    {
        var (workbook, sheet) = Grid();
        var context = new EvaluationContext(workbook);

        // A producer over an open range: its ProbeArray sees the child's refusal and refuses too
        // (the Binary/IF side of Phase 8's split), so the consumer keeps its scalar path, which reaches
        // the producer's own Evaluate → FirstElement → no array → #VALUE!. DynamicArrayTests pins the
        // same outcome for the real FILTER(A:A,A:A>0) as the phase's stated open-range deviation
        // (oracle: #VALUE! plain, 28 CSE).
        var overOpenRange = new Select(
            OpenRangeReference.Create(0, 0, null, null, sheet.Name),
            ArrayAxis.Rows,
            [0]
        );

        await Assert.That(ArrayEvaluation.IsArrayEligible(overOpenRange, context)).IsFalse();
        await Assert.That(ArrayEvaluation.TryEvaluate(overOpenRange, context, out _)).IsFalse();
        await Assert.That(Eval(new Sum([overOpenRange]), workbook)).IsEqualTo(ErrorValue.NotValue);

        // The same, with the refusal answered by the producer itself rather than by a child.
        var refused = new StubProducer(_ => new LiteralArrayOperand(1, 1, 1), Eligible: false);

        await Assert.That(ArrayEvaluation.IsArrayEligible(refused, context)).IsFalse();
        await Assert.That(Eval(new Sum([refused]), workbook)).IsEqualTo(ErrorValue.NotValue);
    }

    // ------------------------------------------------------------------ FirstElement (the @ rule)

    [Test]
    public async Task FirstElement_IsTheTopLeftOfTheProducersArray_OrValueErrorWhenThereIsNone()
    {
        var (workbook, sheet) = Grid();
        var context = new EvaluationContext(workbook);

        // A bare producer in a cell collapses to its TOP-LEFT element — Excel's `@` on an array. Oracle,
        // both modes: =SEQUENCE(2,3) is 1, =SEQUENCE(2,3,7,1) is 7, =SORT(A1:B3,1,-1) is 9 (the descending
        // sort by column 1 orders the rows [3, 1, 2] and the top-left is A3).
        await Assert
            .That(
                Num(
                    ArrayEvaluation
                        .FirstElement(Literal(2, 3, 1, 2, 3, 4, 5, 6), context)
                        .AsObject()
                )
            )
            .IsEqualTo(1.0);
        await Assert
            .That(
                Num(
                    ArrayEvaluation
                        .FirstElement(Literal(2, 3, 7, 8, 9, 10, 11, 12), context)
                        .AsObject()
                )
            )
            .IsEqualTo(7.0);

        var descendingByA = new Select(Range("A1", "B3", sheet), ArrayAxis.Rows, [2, 0, 1]);
        await Assert.That(Num(Eval(descendingByA, workbook))).IsEqualTo(9.0);

        // Under an operator the producer is still an array and the cell still takes the top-left of the
        // RESULT: =SEQUENCE(2,3)*10 is 10 (both modes). FirstElement over the composite says so.
        await Assert
            .That(
                Num(
                    ArrayEvaluation
                        .FirstElement(
                            Multiply(Literal(2, 3, 1, 2, 3, 4, 5, 6), Number(10)),
                            context
                        )
                        .AsObject()
                )
            )
            .IsEqualTo(10.0);

        // No array to take an element from (a refused build) is #VALUE!.
        var refused = new StubProducer(_ => new LiteralArrayOperand(1, 1, 1), Eligible: false);
        await Assert
            .That(ArrayEvaluation.FirstElement(refused, context).AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    // ------------------------------------------------------------------ composition (Phase 10)

    [Test]
    public async Task AProducer_ComposesWithBroadcasting_LikeAnyArrayOperand()
    {
        var (workbook, sheet) = Grid();

        // SEQUENCE(3) is the 3x1 array [1, 2, 3]. Oracle, CSE column (plain is the legacy #VALUE! / 0 of
        // implicit intersection): SUM(SEQUENCE(3)*B1:B3) = 14, COUNT = 3; against the 2x1 H1:H2 the third
        // position is uncovered — SUM(SEQUENCE(3)*H1:H2) = #N/A, COUNT = 2; and the 1x3 SEQUENCE(1,3)
        // against the 3x1 A1:A3 is the 3x3 outer product — SUM = 84, COUNT = 9.
        var column = Literal(3, 1, 1, 2, 3);
        var row = Literal(1, 3, 1, 2, 3);

        await Assert
            .That(Num(Eval(new Sum([Multiply(column, Range("B1", "B3", sheet))]), workbook)))
            .IsEqualTo(14.0);
        await Assert
            .That(Num(Eval(new Count([Multiply(column, Range("B1", "B3", sheet))]), workbook)))
            .IsEqualTo(3.0);
        await Assert
            .That(Eval(new Sum([Multiply(column, Range("H1", "H2", sheet))]), workbook))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Num(Eval(new Count([Multiply(column, Range("H1", "H2", sheet))]), workbook)))
            .IsEqualTo(2.0);
        await Assert
            .That(Num(Eval(new Sum([Multiply(row, Range("A1", "A3", sheet))]), workbook)))
            .IsEqualTo(84.0);
        await Assert
            .That(Num(Eval(new Count([Multiply(row, Range("A1", "A3", sheet))]), workbook)))
            .IsEqualTo(9.0);
    }

    [Test]
    public async Task AOneByOneProducer_BroadcastsLikeAScalar()
    {
        var (workbook, sheet) = Grid();

        // Phase 10 rule 1 applied to a producer's 1x1 result. Oracle, CSE column: SUM(SEQUENCE(1)*A1:A3)
        // = 14, SUM(FILTER(A1,TRUE)*A1:A3) = 70 (a 1x1 SOURCE is still an array result — blocker B1), and
        // SUM(FILTER(A1:A3,A1:A3>100,7)*A1:A3) = 98 (the if_empty singleton fills every position). The
        // singleton is exactly how an EMPTY producer result travels, so this is also why an empty FILTER
        // under an operator errors every element instead of folding to a 0-row extent.
        StubProducer Singleton(double value) =>
            new(_ => new SingletonArrayOperand(ComputedValue.Number(value)));

        var a1a3 = Range("A1", "A3", sheet);

        await Assert
            .That(Num(Eval(new Sum([Multiply(Singleton(1), a1a3)]), workbook)))
            .IsEqualTo(14.0);
        await Assert
            .That(Num(Eval(new Sum([Multiply(Singleton(5), a1a3)]), workbook)))
            .IsEqualTo(70.0);
        await Assert
            .That(Num(Eval(new Sum([Multiply(Singleton(7), a1a3)]), workbook)))
            .IsEqualTo(98.0);
        await Assert
            .That(Num(Eval(new Count([Multiply(Singleton(7), a1a3)]), workbook)))
            .IsEqualTo(3.0);

        // A scalar SOURCE under a selection: Select wraps it (B1), so SUM(SORT(A1)) = 5 and ROWS = 1
        // (both modes, DynamicArrayTests) are what the contract yields for a cell.
        var overCell = new Select(Cell("A1", sheet), ArrayAxis.Rows, [0]);
        await Assert.That(Num(Eval(new Sum([overCell]), workbook))).IsEqualTo(5.0);
        await Assert
            .That(
                ArrayEvaluation.TryEvaluate(overCell, new EvaluationContext(workbook), out var one)
            )
            .IsTrue();
        await Assert.That(one.Rows).IsEqualTo(1);
        await Assert.That(one.Columns).IsEqualTo(1);
    }

    [Test]
    public async Task AProducer_UnderAnIf_ALiftedFunction_AndAUnary_IsReachedByRecursion()
    {
        var (workbook, _) = Grid();

        // Composition needs no code: the existing arms recurse into the producer arm. Oracle, both modes:
        // SUM(IF(SEQUENCE(3)>1,SEQUENCE(3),0)) = 5, SUM(LEN(SEQUENCE(3)*10)) = 6, SUM(-SEQUENCE(3)) = -6.
        StubProducer Sequence3() => Literal(3, 1, 1, 2, 3);

        var underIf = new If([GreaterThan(Sequence3(), Number(1)), Sequence3(), Number(0)]);
        var underLen = new Len([Multiply(Sequence3(), Number(10))]);
        var underNegate = Negate(Sequence3());

        await Assert.That(Num(Eval(new Sum([underIf]), workbook))).IsEqualTo(5.0);
        await Assert.That(Num(Eval(new Sum([underLen]), workbook))).IsEqualTo(6.0);
        await Assert.That(Num(Eval(new Sum([underNegate]), workbook))).IsEqualTo(-6.0);
    }

    [Test]
    public async Task AProducer_InsideAProducer_ComposesThroughTheRecursiveBuilder()
    {
        var (workbook, sheet) = Grid();

        // Oracle, both modes: SUM(FILTER(SEQUENCE(5),SEQUENCE(5)>2)) = 12 — rows 3..5 of the 5x1
        // sequence. Select's child here is another producer, reached through TryBuildOperand's recursion.
        var lastThree = new Select(Literal(5, 1, 1, 2, 3, 4, 5), ArrayAxis.Rows, [2, 3, 4]);

        await Assert.That(Num(Eval(new Sum([lastThree]), workbook))).IsEqualTo(12.0);

        // And a selection of a selection: the descending-by-A permutation of A1:B3, then its first two
        // rows — [9, 3], [5, 1] — is what INDEX(SORT(A1:B3,1,-1),k,c) reads for k ≤ 2 (oracle, both
        // modes: (1,2) = 3, (2,1) = 5).
        var sorted = new Select(Range("A1", "B3", sheet), ArrayAxis.Rows, [2, 0, 1]);
        var topTwo = new Select(sorted, ArrayAxis.Rows, [0, 1]);

        await Assert
            .That(Num(Eval(new Index([topTwo, Number(1), Number(2)]), workbook)))
            .IsEqualTo(3.0);
        await Assert
            .That(Num(Eval(new Index([topTwo, Number(2), Number(1)]), workbook)))
            .IsEqualTo(5.0);
        await Assert.That(Num(Eval(new Count([topTwo]), workbook))).IsEqualTo(4.0);
    }

    // ------------------------------------------------------------------ AxisSelectionOperand

    [Test]
    public async Task AxisSelection_MapsTheChosenAxis_AndReadsTheSourceAtTheSourcesOwnExtent()
    {
        var (workbook, sheet) = Grid();
        var context = new EvaluationContext(workbook);

        // Rows [3, 1, 2] of A1:B3 = [[5,1],[0,2],[9,3]] is SORT(A1:B3,1,-1) = [[9,3],[5,1],[0,2]].
        // Oracle: INDEX(SORT(A1:B3,1,-1),1,2) = 3 and (2,1) = 5 (both modes); under broadcasting,
        // SUM(SORT(A1:B3,1,-1)*B1:B3) = 30 (CSE; the 3x1 column repeats across both columns) and against
        // the 2x1 H1:H2 the third row is uncovered — SUM = #N/A, COUNT = 4 (CSE).
        var byRows = new Select(Range("A1", "B3", sheet), ArrayAxis.Rows, [2, 0, 1]);

        await Assert.That(ArrayEvaluation.TryEvaluate(byRows, context, out var rows)).IsTrue();
        await Assert.That(rows.Rows).IsEqualTo(3);
        await Assert.That(rows.Columns).IsEqualTo(2);
        await Assert.That(NumberAt(rows, 0)).IsEqualTo(9.0);
        await Assert.That(NumberAt(rows, 1)).IsEqualTo(3.0);
        await Assert.That(NumberAt(rows, 2)).IsEqualTo(5.0);
        await Assert.That(NumberAt(rows, 3)).IsEqualTo(1.0);
        await Assert.That(NumberAt(rows, 4)).IsEqualTo(0.0);
        await Assert.That(NumberAt(rows, 5)).IsEqualTo(2.0);
        await Assert
            .That(Num(Eval(new Sum([Multiply(byRows, Range("B1", "B3", sheet))]), workbook)))
            .IsEqualTo(30.0);
        await Assert
            .That(Eval(new Sum([Multiply(byRows, Range("H1", "H2", sheet))]), workbook))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Num(Eval(new Count([Multiply(byRows, Range("H1", "H2", sheet))]), workbook)))
            .IsEqualTo(4.0);

        // Columns [2, 1] of A1:B3 is SORT(A1:B3,1,1,TRUE): the columns swap, so (1,1) = 1 and (1,2) = 5
        // (both modes, DynamicArrayTests). The selected axis is the COLUMN one this time.
        var byColumns = new Select(Range("A1", "B3", sheet), ArrayAxis.Columns, [1, 0]);

        await Assert
            .That(ArrayEvaluation.TryEvaluate(byColumns, context, out var columns))
            .IsTrue();
        await Assert.That(columns.Rows).IsEqualTo(3);
        await Assert.That(columns.Columns).IsEqualTo(2);
        await Assert.That(NumberAt(columns, 0)).IsEqualTo(1.0);
        await Assert.That(NumberAt(columns, 1)).IsEqualTo(5.0);
        await Assert.That(NumberAt(columns, 5)).IsEqualTo(9.0);

        // A 2x1 selection against a 3x1 range: FILTER(A1:A3,A1:A3>0) keeps rows [1, 3] and
        // SUM(FILTER(A1:A3,A1:A3>0)*B1:B3) = #N/A with COUNT 2 (CSE) — the selection's OWN extent is what
        // it projects, not its source's.
        var kept = new Select(Range("A1", "A3", sheet), ArrayAxis.Rows, [0, 2]);

        await Assert
            .That(Eval(new Sum([Multiply(kept, Range("B1", "B3", sheet))]), workbook))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Num(Eval(new Count([Multiply(kept, Range("B1", "B3", sheet))]), workbook)))
            .IsEqualTo(2.0);

        // The source is read at ITS extent, whatever it is — a composite source included (engine
        // mechanism, no oracle formula): rows [3, 1] of A1:A3*2 are 18 and 10.
        var ofComposite = new Select(
            Multiply(Range("A1", "A3", sheet), Number(2)),
            ArrayAxis.Rows,
            [2, 0]
        );

        await Assert
            .That(ArrayEvaluation.TryEvaluate(ofComposite, context, out var doubled))
            .IsTrue();
        await Assert.That(doubled.Length).IsEqualTo(2);
        await Assert.That(NumberAt(doubled, 0)).IsEqualTo(18.0);
        await Assert.That(NumberAt(doubled, 1)).IsEqualTo(10.0);
    }

    // ------------------------------------------------------------------ the invariant

    [Test]
    public async Task AnEmptyOrScalarSelection_IsRejectedAtConstruction_TheInvariantIsEnforcedNotAspirational()
    {
        var (workbook, sheet) = Grid();
        var context = new EvaluationContext(workbook);

        await Assert
            .That(
                ArrayEvaluation.TryBuildOperand(Range("A1", "B3", sheet), context, out var source)
            )
            .IsTrue();

        // Nothing selected is a 0-extent array, which the invariant forbids: the producer must hand back a
        // 1x1 SingletonArrayOperand carrying the error instead. The guard names the rule.
        await Assert
            .That(() => new AxisSelectionOperand(source, ArrayAxis.Rows, []))
            .Throws<InvalidOperationException>()
            .WithMessageContaining("Rows >= 1 && Columns >= 1");
        await Assert
            .That(() => new AxisSelectionOperand(source, ArrayAxis.Columns, []))
            .Throws<InvalidOperationException>();

        // A scalar SOURCE reports 0x0 (ScalarOperand), so a selection over it would be 0-extent on the
        // other axis: blocker B1 says wrap it in a SingletonArrayOperand first.
        await Assert
            .That(() =>
                new AxisSelectionOperand(
                    new ScalarOperand(ComputedValue.Number(5)),
                    ArrayAxis.Rows,
                    [0]
                )
            )
            .Throws<InvalidOperationException>();

        // A selection past the source's axis is a producer bug, caught at build rather than as a stray
        // IndexOutOfRangeException per element.
        await Assert
            .That(() => new AxisSelectionOperand(source, ArrayAxis.Rows, [3]))
            .Throws<InvalidOperationException>();
    }

    [Test]
    public async Task AZeroExtentOperand_MakesEveryConsumerSilentlyWrong_WhichIsWhatTheInvariantPrevents()
    {
        var (workbook, _) = Grid();

        // WHAT BREAKS if the invariant is violated, measured on this tree: ArrayStream.Length is
        // Rows * Columns with no emptiness channel, so a 0x1 operand enumerates NOTHING and every
        // consumer answers as if the argument were an empty range — SUM 0, COUNT 0, AVERAGE #DIV/0! —
        // where the oracle answers #CALC! for the empty FILTER (SUM(FILTER(A1:A3,A1:A3>100)), both
        // modes; DynamicArrayTests). A silent 0 in place of an error is the exact bug class the
        // invariant exists to make impossible; this test bypasses ArrayShaping on purpose to show it.
        var empty = new StubProducer(_ => new LiteralArrayOperand(0, 1));

        await Assert.That(Num(Eval(new Sum([empty]), workbook))).IsEqualTo(0.0);
        await Assert.That(Num(Eval(new Count([empty]), workbook))).IsEqualTo(0.0);
        await Assert.That(Eval(new Average([empty]), workbook)).IsEqualTo(ErrorValue.DivByZero);
    }
}
