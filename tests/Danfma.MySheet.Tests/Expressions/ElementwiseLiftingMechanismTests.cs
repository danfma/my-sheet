using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;
using StringValue = Danfma.MySheet.Expressions.StringValue; // TUnit also defines a StringValue

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Unit tests for the MECHANISM behind Phase 8's elementwise lifting (the acceptance pins live in
/// <see cref="ElementwiseLiftingTests"/>): the mutable <see cref="ScratchLiteral"/> slot, the two new operands
/// (<see cref="UnaryOperand"/>, <see cref="LiftedFunctionOperand"/>), and the <c>Probe</c>/<c>TryBuildOperand</c>
/// arms in <see cref="ArrayEvaluation"/> that route a <c>Negate</c>/<c>Percent</c> or an
/// <see cref="ArrayLifting.Elementwise"/> built-in to them. The invariants pinned here are the ones the
/// design leans on: the node is built ONCE per evaluation and re-evaluated per element over rebound slots; an
/// omitted optional argument keeps its literal <see cref="BlankValue"/> slot (verifier correction B1); an
/// argument whose build refuses turns the whole function into an OPAQUE scalar, never a hard refusal (M1);
/// the probe never evaluates; and a scalar argument — a volatile in particular — is evaluated exactly once.
/// </summary>
public class ElementwiseLiftingMechanismTests
{
    private static (Workbook Workbook, Sheet Sheet) Sheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return (workbook, sheet);
    }

    private static Expression Parse(string formula, Sheet sheet) =>
        ExpressionParser.Parse(formula, sheet);

    private static object? Evaluate(string formula, Sheet sheet, Workbook workbook) =>
        Parse(formula, sheet).Evaluate(workbook).AsObject();

    private static ArrayEvaluationResult Array(string formula, Sheet sheet, Workbook workbook)
    {
        var context = new EvaluationContext(workbook);
        var built = ArrayEvaluation.TryEvaluate(Parse(formula, sheet), context, out var result);

        return built ? result : throw new InvalidOperationException($"{formula} is not an array");
    }

    private static double NumberAt(ArrayEvaluationResult result, int index)
    {
        result.Values[index].TryGetNumber(out var value);
        return value;
    }

    private static string? TextAt(ArrayEvaluationResult result, int index) =>
        result.Values[index].AsString();

    private static Error? ErrorAt(ArrayEvaluationResult result, int index) =>
        result.Values[index].TryGetError(out var error) ? error : null;

    // --- ScratchLiteral ---

    [Test]
    public async Task ScratchLiteral_EvaluatesWhateverValueItCurrentlyHolds()
    {
        var (workbook, _) = Sheet();
        var context = new EvaluationContext(workbook);
        var slot = new ScratchLiteral { Value = ComputedValue.Number(2) };

        await Assert.That(slot.Evaluate(context).AsDouble()).IsEqualTo(2.0);

        // The one mutable node in the engine: rebinding the slot changes what the SAME node evaluates to,
        // which is what lets a lifted node be built once and driven across every element.
        slot.Value = ComputedValue.Text("abc");

        await Assert.That(slot.Evaluate(context).AsString()).IsEqualTo("abc");
    }

    // --- UnaryOperand ---

    [Test]
    public async Task UnaryOperand_NegateAndPercent_OverTheInnerElement()
    {
        // A 3x1 position vector [1,2,3] as the inner array — no sheet needed.
        var inner = new PositionNumbersOperand(1, PositionAxis.Row, 3, 1);

        var negate = new UnaryOperand(UnaryOperator.Negate, inner, 3, 1);
        await Assert.That(negate.IsArray).IsTrue();
        await Assert.That(negate.Rows).IsEqualTo(3);
        await Assert.That(negate.Columns).IsEqualTo(1);
        await Assert.That(negate.At(0, 3, 1).AsDouble()).IsEqualTo(-1.0);
        await Assert.That(negate.At(2, 3, 1).AsDouble()).IsEqualTo(-3.0);

        var percent = new UnaryOperand(UnaryOperator.Percent, inner, 3, 1);
        await Assert.That(percent.At(1, 3, 1).AsDouble()).IsEqualTo(0.02);
    }

    [Test]
    public async Task UnaryOperand_TextElement_IsAValueError()
    {
        var operand = new UnaryOperand(
            UnaryOperator.Negate,
            new ScalarOperand(ComputedValue.Text("x")),
            2,
            1
        );

        operand.At(0, 2, 1).TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.Value);
    }

    [Test]
    public async Task UnaryOperand_ShapeMismatch_IsAValueError()
    {
        // The same guard every array operand carries: asked for a shape other than its own, every element is
        // #VALUE! (Excel's dimension-mismatch rule).
        var operand = new UnaryOperand(
            UnaryOperator.Negate,
            new PositionNumbersOperand(1, PositionAxis.Row, 3, 1),
            3,
            1
        );

        operand.At(0, 2, 1).TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.Value);
    }

    // --- LiftedFunctionOperand ---

    [Test]
    public async Task LiftedFunctionOperand_BuildsTheNodeOnce_AndRebindsTheSlotsPerElement()
    {
        var (workbook, _) = Sheet();
        var context = new EvaluationContext(workbook);
        var creations = 0;

        // Elements [1,2,3]; the "function" is a Negate over slot 0, so the answers are -1,-2,-3.
        var operand = new LiftedFunctionOperand(
            slots =>
            {
                creations++;
                return new UnaryOperation(UnaryOperator.Negate, slots[0]);
            },
            [Number(0)],
            [new PositionNumbersOperand(1, PositionAxis.Row, 3, 1)],
            context,
            3,
            1
        );

        await Assert.That(operand.IsArray).IsTrue();
        await Assert.That(operand.Rows).IsEqualTo(3);
        await Assert.That(operand.Columns).IsEqualTo(1);
        await Assert.That(operand.At(0, 3, 1).AsDouble()).IsEqualTo(-1.0);
        await Assert.That(operand.At(1, 3, 1).AsDouble()).IsEqualTo(-2.0);
        await Assert.That(operand.At(2, 3, 1).AsDouble()).IsEqualTo(-3.0);

        // The whole point of the design: one node for the whole array, not one per element.
        await Assert.That(creations).IsEqualTo(1);
    }

    [Test]
    public async Task LiftedFunctionOperand_ShapeMismatch_IsAValueError()
    {
        var (workbook, _) = Sheet();
        var context = new EvaluationContext(workbook);

        var operand = new LiftedFunctionOperand(
            slots => new UnaryOperation(UnaryOperator.Negate, slots[0]),
            [Number(0)],
            [new PositionNumbersOperand(1, PositionAxis.Row, 3, 1)],
            context,
            3,
            1
        );

        operand.At(0, 1, 1).TryGetError(out var error);
        await Assert.That(error).IsEqualTo(Error.Value);
    }

    [Test]
    public async Task LiftedFunctionOperand_KeepsAnOmittedArgumentSlot_AsTheBlankValueLiteral()
    {
        // Verifier correction B1. Ten of the lifted built-ins detect an OMITTED optional argument by
        // pattern-matching the literal BlankValue node the parser leaves in its slot (`Arguments[1] is not
        // BlankValue`). A ScratchLiteral is not a BlankValue, so such a slot must be handed through untouched:
        // only the non-blank slots become scratch slots.
        var (workbook, _) = Sheet();
        var context = new EvaluationContext(workbook);
        Expression[]? seen = null;

        _ = new LiftedFunctionOperand(
            slots =>
            {
                seen = slots;
                return new UnaryOperation(UnaryOperator.Negate, slots[0]);
            },
            [Number(0), BlankValue.Instance, Number(0)],
            [
                new PositionNumbersOperand(1, PositionAxis.Row, 3, 1),
                new ScalarOperand(ComputedValue.Blank),
                new ScalarOperand(ComputedValue.Number(1)),
            ],
            context,
            3,
            1
        );

        await Assert.That(seen).IsNotNull();
        await Assert.That(seen![0]).IsTypeOf<ScratchLiteral>();
        await Assert.That(ReferenceEquals(seen[1], BlankValue.Instance)).IsTrue();
        await Assert.That(seen[2]).IsTypeOf<ScratchLiteral>();
    }

    [Test]
    public async Task OmittedArgument_ThroughTheParser_KeepsTheFunctionsDefault()
    {
        // B1 end to end, on the three shapes the verifier measured. FIXED(x,,TRUE): the omitted decimals keep
        // the default of 2 (measured broken value before the correction: "1235" — decimals coerced to 0).
        var (workbook, sheet) = Sheet();
        sheet["A1"] = new NumberValue(1234.5);
        sheet["A2"] = new NumberValue(2.345);
        sheet["A3"] = new NumberValue(0);

        var fixedText = Array("=FIXED(A1:A3,,TRUE)", sheet, workbook);
        await Assert.That(TextAt(fixedText, 0)).IsEqualTo("1234.50");
        await Assert.That(TextAt(fixedText, 1)).IsEqualTo("2.35");
        await Assert.That(TextAt(fixedText, 2)).IsEqualTo("0.00");

        // Same for the scalar shape, so the lifted answer is the scalar answer element by element.
        await Assert.That(Evaluate("=FIXED(A1,,TRUE)", sheet, workbook)).IsEqualTo("1234.50");

        var dollar = Array("=DOLLAR(A1:A3,)", sheet, workbook);
        await Assert.That(TextAt(dollar, 0)).IsEqualTo("$1,234.50");
        await Assert.That(TextAt(dollar, 1)).IsEqualTo("$2.35");

        sheet["C1"] = new StringValue("a-b-c");
        sheet["C2"] = new StringValue("d-e");
        sheet["C3"] = new StringValue("f");

        // TEXTBEFORE(text, "-", [instance_num omitted], match_mode 1): the omitted instance is the FIRST.
        var before = Array("=TEXTBEFORE(C1:C3,\"-\",,1)", sheet, workbook);
        await Assert.That(TextAt(before, 0)).IsEqualTo("a");
        await Assert.That(TextAt(before, 1)).IsEqualTo("d");
    }

    // --- The Probe arms ---

    [Test]
    public async Task Probe_LiftsNegateAndPercent_AndNeverPlus()
    {
        var (workbook, sheet) = Sheet();
        var context = new EvaluationContext(workbook);

        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=-(A1:A3)", sheet), context))
            .IsTrue();
        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=A1:A3%", sheet), context))
            .IsTrue();

        // A Negate over a scalar is a scalar; unary Plus is the reference-preserving no-op and stays opaque.
        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=-(5)", sheet), context))
            .IsFalse();
        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=+A1:A3", sheet), context))
            .IsFalse();
    }

    [Test]
    public async Task Probe_LiftsAnElementwiseBuiltIn_ExactlyWhenAnArgumentIsAnArray()
    {
        var (workbook, sheet) = Sheet();
        var context = new EvaluationContext(workbook);

        bool Eligible(string formula) =>
            ArrayEvaluation.IsArrayEligible(Parse(formula, sheet), context);

        await Assert.That(Eligible("=LEN(A1:A3)")).IsTrue();
        await Assert.That(Eligible("=LEN(TRIM(A1:A3))")).IsTrue();
        await Assert.That(Eligible("=ROUND(1,A1:A3)")).IsTrue();
        await Assert.That(Eligible("=MID(A1:A3,B1:B3,C1:C3)")).IsTrue();

        // Every argument scalar → the function is a scalar (evaluated once, broadcast if nested).
        await Assert.That(Eligible("=LEN(\"abc\")")).IsFalse();
        await Assert.That(Eligible("=ROUND(1,2)")).IsFalse();

        // A zero-argument entry never lifts; a Consumes entry never lifts, whatever its argument.
        await Assert.That(Eligible("=PI()")).IsFalse();
        await Assert.That(Eligible("=SUM(A1:A3)")).IsFalse();
        await Assert.That(Eligible("=ROW(A1:A3)")).IsTrue(); // its own dedicated arm, not the lift
    }

    [Test]
    public async Task Probe_ARefusedArgument_MakesTheFunctionAnOpaqueScalar_NotARefusal()
    {
        // Verifier correction M1. An open range refuses the mini-CSE (the cost guard); through the new arms
        // that refusal is TOLERATED as the opaque-scalar answer, exactly as the pre-phase `default` arm
        // tolerated it, so an enclosing array expression keeps its eligibility.
        var (workbook, sheet) = Sheet();
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        var context = new EvaluationContext(workbook);

        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=LEN(A:A)", sheet), context))
            .IsFalse();
        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=-(A:A)", sheet), context))
            .IsFalse();
        await Assert
            .That(ArrayEvaluation.IsArrayEligible(Parse("=IF(A1:A3>0,1,LEN(B:B))", sheet), context))
            .IsTrue();

        // The regression the correction removes: today's answer, kept.
        await Assert.That(Evaluate("=SUM(IF(A1:A3>0,1,LEN(B:B)))", sheet, workbook)).IsEqualTo(3.0);

        // And the risk section's answer, kept: the opaque scalar over an open range is #VALUE!.
        await Assert
            .That(Evaluate("=SUM(LEN(A:A))", sheet, workbook))
            .IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Evaluate("=SUM(ROUND(A1:A3,LEN(B:B)))", sheet, workbook))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Probe_NeverEvaluates()
    {
        var (workbook, sheet) = Sheet();
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);

        var draws = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++draws);
        var context = new EvaluationContext(workbook);

        int DrawsOfProbe(string formula)
        {
            draws = 0;
            ArrayEvaluation.IsArrayEligible(Parse(formula, sheet), context);
            return draws;
        }

        await Assert.That(DrawsOfProbe("=ROUND(A1:A3,TICK())")).IsEqualTo(0);
        await Assert.That(DrawsOfProbe("=IFERROR(TICK(),A1:A3)")).IsEqualTo(0);
        await Assert.That(DrawsOfProbe("=-(A1:A3*TICK())")).IsEqualTo(0);
        await Assert.That(DrawsOfProbe("=LEN(TRIM(A1:A3))")).IsEqualTo(0);
    }

    // --- The build arms ---

    [Test]
    public async Task Build_AScalarArgument_IsEvaluatedOnce_AndBroadcast()
    {
        // The volatile-once invariant: a scalar argument becomes a ScalarOperand at build time and is
        // broadcast, so a volatile (here a counting custom function) draws exactly once per evaluation —
        // the same count the scalar path and SUM(IF(A1:A3>0,TICK(),0)) give today.
        var (workbook, sheet) = Sheet();
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);

        var draws = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++draws);

        int DrawsOf(string formula)
        {
            draws = 0;
            Parse(formula, sheet).Evaluate(workbook);
            return draws;
        }

        await Assert.That(DrawsOf("=SUM(ROUND(A1:A3,TICK()))")).IsEqualTo(1);
        await Assert.That(DrawsOf("=SUM(IFERROR(A1:A3,TICK()))")).IsEqualTo(1);
        await Assert.That(DrawsOf("=SUM(ROUND(A1:A3*TICK(),0))")).IsEqualTo(1);
        await Assert.That(DrawsOf("=SUM(-(A1:A3*TICK()))")).IsEqualTo(1);

        // A lifted-CLASS node whose arguments are all scalars, NESTED inside an array expression: it is an
        // opaque scalar there, and must still be evaluated exactly once (not once to discover it is scalar
        // and once more for its value).
        await Assert.That(DrawsOf("=SUM(A1:A3+ROUND(TICK(),0))")).IsEqualTo(1);
        await Assert.That(DrawsOf("=SUM(A1:A3+(-TICK()))")).IsEqualTo(1);
        await Assert.That(DrawsOf("=SUM(A1:A3+ROUND(TICK(),LEN(B:B)))")).IsEqualTo(1);

        // The pre-phase reference count, unchanged.
        await Assert.That(DrawsOf("=SUM(IF(A1:A3>0,TICK(),0))")).IsEqualTo(1);
    }

    [Test]
    public async Task Build_NAryShape_FoldsLikeTheBinaryRule()
    {
        var (workbook, sheet) = Sheet();
        sheet["A1"] = new StringValue("abcdef");
        sheet["A2"] = new StringValue("ghijkl");
        sheet["A3"] = new StringValue("mnopqr");
        sheet["B1"] = new NumberValue(1);
        sheet["B2"] = new NumberValue(2);
        sheet["B3"] = new NumberValue(3);
        sheet["C1"] = new NumberValue(1);
        sheet["C2"] = new NumberValue(2);
        sheet["C3"] = new NumberValue(3);

        // Three same-shaped arrays zip position by position.
        var mid = Array("=MID(A1:A3,B1:B3,C1:C3)", sheet, workbook);
        await Assert.That(mid.Rows).IsEqualTo(3);
        await Assert.That(mid.Columns).IsEqualTo(1);
        await Assert.That(TextAt(mid, 0)).IsEqualTo("a");
        await Assert.That(TextAt(mid, 1)).IsEqualTo("hi");
        await Assert.That(TextAt(mid, 2)).IsEqualTo("opq");

        // A scalar in the middle broadcasts; the shape comes from the arrays alone.
        var broadcast = Array("=MID(A1:A3,2,C1:C3)", sheet, workbook);
        await Assert.That(TextAt(broadcast, 0)).IsEqualTo("b");
        await Assert.That(TextAt(broadcast, 2)).IsEqualTo("nop");

        // Mismatched arrays (3x1 against 2x1) take the LARGER extent — ShapeFold through Broadcasting.Axis,
        // the same rule as BinaryOperand's, not a second one — and the shorter argument's own At() answers
        // #N/A at the one position it does not cover, so MID reads a real start at rows 1-2 and the error at
        // row 3. FLIPPED by Phase 10 (verifier correction M2) from #VALUE! at elements 0 and 2; measured on
        // Aspose.Cells 26.6.0 (2026-09-09, CSE column, the verifier's own fixture): the third element is
        // #N/A and the first two are the MID of their rows.
        var mismatch = Array("=MID(A1:A3,B1:B2,1)", sheet, workbook);
        await Assert.That(mismatch.Rows).IsEqualTo(3);
        await Assert.That(mismatch.Columns).IsEqualTo(1);
        await Assert.That(TextAt(mismatch, 0)).IsEqualTo("a");
        await Assert.That(TextAt(mismatch, 1)).IsEqualTo("h");
        await Assert.That(ErrorAt(mismatch, 2)).IsEqualTo(Error.NA);
    }

    [Test]
    public async Task BroadcastArgument_ErrorConsumingBody_SeesNoMarker()
    {
        // IFERROR(A1:C3, E1:E3) = 9, and the REASON matters because two earlier versions of this comment got
        // it wrong in opposite directions. IFERROR SHORT-CIRCUITS on its first argument alone: it evaluates
        // argument 1 only when argument 0 is an error. Here argument 0 is A1:C3, which never errors at any
        // position, so argument 1 is never read — before Phase 10 or after it. The lift still computes the
        // second argument's element into its scratch slot (a #VALUE! marker before this phase, a broadcast
        // number now), and IFERROR simply never looks at it. So the count is 9 for one reason only: nine
        // non-error elements on the left. Aspose.Cells 26.6.0 answers 9 as well (2026-09-09, CSE column), but
        // that agreement is not what this test is about. The companion where IFERROR DOES see a real error is
        // VectorBroadcastingTests' SUM(IFERROR(A1:C3*H1:H2,0)) = 36, whose left side has #N/A in the tail.
        var (workbook, sheet) = Sheet();
        foreach (var column in new[] { "A", "B", "C", "E" })
        {
            for (var row = 1; row <= 3; row++)
            {
                sheet[$"{column}{row}"] = new NumberValue(row);
            }
        }

        await Assert.That(Evaluate("=COUNT(IFERROR(A1:C3,E1:E3))", sheet, workbook)).IsEqualTo(9.0);
    }

    [Test]
    public async Task ShapeMismatch_ErrorConsumingBody_SeesNoMarkerOnceTheColumnBroadcasts()
    {
        // The companion of the pin above. LEFT(D7:F9,E6:E8) is a 3x3 against a 3x1 whose cells are BLANK,
        // so with the column broadcast every element is LEFT(x, 0) = "" — there is no error left for
        // IFERROR to recover and LEN sums to 0. FLIPPED by Phase 10 (verifier correction M2) from 18: before
        // it, the 3x1 side answered nine #VALUE! markers that IFERROR turned into "zz". Aspose.Cells 26.6.0
        // answers 0 (measured 2026-09-09, CSE column).
        var (textbook, textSheet) = Sheet();
        textSheet["D7"] = new StringValue("abc");
        textSheet["D8"] = new StringValue("def");
        textSheet["E9"] = new StringValue(" ");

        await Assert
            .That(Evaluate("=SUM(LEN(IFERROR(LEFT(D7:F9,E6:E8),\"zz\")))", textSheet, textbook))
            .IsEqualTo(0.0);
    }

    [Test]
    public async Task Build_UnaryOverAScalar_IsTheScalarAnswer()
    {
        // Nested inside an array expression, -(scalar) is an opaque scalar evaluated once and broadcast.
        var (workbook, sheet) = Sheet();
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);

        var result = Array("=A1:A3+(-(2))", sheet, workbook);
        await Assert.That(NumberAt(result, 0)).IsEqualTo(-1.0);
        await Assert.That(NumberAt(result, 2)).IsEqualTo(1.0);

        var percent = Array("=A1:A3+(50%)", sheet, workbook);
        await Assert.That(NumberAt(percent, 0)).IsEqualTo(1.5);
    }
}
