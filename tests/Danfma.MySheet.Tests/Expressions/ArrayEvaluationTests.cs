using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Direct unit tests for the internal <see cref="ArrayEvaluation"/> element-wise evaluator (Phase A of the
/// mini-CSE plan). No production code consumes the evaluator yet; these drive the vector semantics in
/// isolation. Golden values come from the plan's oracle (Excel implicit-intersection / CSE semantics).
/// </summary>
public class ArrayEvaluationTests
{
    // B2..B5 = Hide/Show/Hide/Show — the K1 repro fixture.
    private static (Workbook Workbook, Sheet Sheet) ShowHideGrid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["B2"] = String("Hide");
        sheet["B3"] = String("Show");
        sheet["B4"] = String("Hide");
        sheet["B5"] = String("Show");

        return (workbook, sheet);
    }

    private static BinaryOperation Equal(Expression left, Expression right) =>
        new(BinaryOperator.Equal, left, right);

    private static bool BoolAt(ArrayEvaluationResult result, int index)
    {
        result.Values[index].TryGetBoolean(out var value);
        return value;
    }

    private static double NumberAt(ArrayEvaluationResult result, int index)
    {
        result.Values[index].TryGetNumber(out var value);
        return value;
    }

    [Test]
    public async Task RangeComparison_ProducesBooleanVector()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // B2:B5="Show" → [FALSE, TRUE, FALSE, TRUE]
        var node = Equal(Range("B2", "B5", sheet), String("Show"));

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Rows).IsEqualTo(4);
        await Assert.That(result.Columns).IsEqualTo(1);
        await Assert.That(result.Length).IsEqualTo(4);
        await Assert.That(result.Values[0].Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BoolAt(result, 0)).IsFalse();
        await Assert.That(BoolAt(result, 1)).IsTrue();
        await Assert.That(BoolAt(result, 2)).IsFalse();
        await Assert.That(BoolAt(result, 3)).IsTrue();
    }

    [Test]
    public async Task IfWithArrayCondition_ThenElseScalars_Broadcasts()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // IF(B2:B5="Show", 1, 0) → [0, 1, 0, 1]
        var node = new If([Equal(Range("B2", "B5", sheet), String("Show")), Number(1), Number(0)]);

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Length).IsEqualTo(4);
        await Assert.That(NumberAt(result, 0)).IsEqualTo(0.0);
        await Assert.That(NumberAt(result, 1)).IsEqualTo(1.0);
        await Assert.That(NumberAt(result, 2)).IsEqualTo(0.0);
        await Assert.That(NumberAt(result, 3)).IsEqualTo(1.0);
    }

    [Test]
    public async Task IfWithoutElse_YieldsLogicalFalseWhereConditionFalse()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // IF(B2:B5="Show", ROW(B2:B5)) → [FALSE, 3, FALSE, 5]
        var node = new If(
            [Equal(Range("B2", "B5", sheet), String("Show")), new Row([Range("B2", "B5", sheet)])]
        );

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Length).IsEqualTo(4);

        await Assert.That(result.Values[0].Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BoolAt(result, 0)).IsFalse();
        await Assert.That(NumberAt(result, 1)).IsEqualTo(3.0);
        await Assert.That(result.Values[2].Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BoolAt(result, 2)).IsFalse();
        await Assert.That(NumberAt(result, 3)).IsEqualTo(5.0);
    }

    [Test]
    public async Task RowOfRange_ProducesRowNumberVector()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // ROW(B2:B5) → [2, 3, 4, 5]
        var node = new Row([Range("B2", "B5", sheet)]);

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Length).IsEqualTo(4);
        await Assert.That(NumberAt(result, 0)).IsEqualTo(2.0);
        await Assert.That(NumberAt(result, 1)).IsEqualTo(3.0);
        await Assert.That(NumberAt(result, 2)).IsEqualTo(4.0);
        await Assert.That(NumberAt(result, 3)).IsEqualTo(5.0);
    }

    [Test]
    public async Task Arithmetic_ArrayTimesScalar_Broadcasts()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // ROW(B2:B5) * 2 → [4, 6, 8, 10]
        var node = new BinaryOperation(
            BinaryOperator.Multiply,
            new Row([Range("B2", "B5", sheet)]),
            Number(2)
        );

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Length).IsEqualTo(4);
        await Assert.That(NumberAt(result, 0)).IsEqualTo(4.0);
        await Assert.That(NumberAt(result, 1)).IsEqualTo(6.0);
        await Assert.That(NumberAt(result, 2)).IsEqualTo(8.0);
        await Assert.That(NumberAt(result, 3)).IsEqualTo(10.0);
    }

    [Test]
    public async Task NestedIf_Bh25Shape_FalseWhereAnyConditionFails()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // IF(B2:B5="Show", IF(ROW(B2:B5) > 3, ROW(B2:B5)))
        // condition [F,T,F,T]; inner keeps ROW only where ROW>3 → [_, F(row3), _, 5]
        // net → [FALSE, FALSE, FALSE, 5]
        var range = Range("B2", "B5", sheet);
        var inner = new If(
            [GreaterThan(new Row([range]), Number(3)), new Row([range])]
        );
        var node = new If([Equal(range, String("Show")), inner]);

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Length).IsEqualTo(4);
        await Assert.That(result.Values[0].Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BoolAt(result, 0)).IsFalse();
        // B3 matches "Show" but ROW(3) is not > 3 → inner IF-without-else yields FALSE.
        await Assert.That(result.Values[1].Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BoolAt(result, 1)).IsFalse();
        await Assert.That(result.Values[2].Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(BoolAt(result, 2)).IsFalse();
        await Assert.That(NumberAt(result, 3)).IsEqualTo(5.0);
    }

    [Test]
    public async Task DimensionMismatch_YieldsValueErrorPerElement()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Number(3);
        sheet["C1"] = Number(10);
        sheet["C2"] = Number(20);
        var context = new EvaluationContext(workbook);

        // A1:A3 + C1:C2 → dims mismatch (3x1 vs 2x1) → every element #VALUE!
        var node = new BinaryOperation(
            BinaryOperator.Add,
            Range("A1", "A3", sheet),
            Range("C1", "C2", sheet)
        );

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        foreach (var value in result.Values)
        {
            await Assert.That(value.TryGetError(out var error)).IsTrue();
            await Assert.That(error).IsEqualTo(Error.Value);
        }
    }

    [Test]
    public async Task ErrorInCell_IsPreservedAtPosition()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = Number(1);
        sheet["A2"] = Divide(Number(1), Number(0)); // #DIV/0!
        sheet["A3"] = Number(3);
        var context = new EvaluationContext(workbook);

        // A1:A3 * 2 → [2, #DIV/0!, 6]
        var node = new BinaryOperation(BinaryOperator.Multiply, Range("A1", "A3", sheet), Number(2));

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(NumberAt(result, 0)).IsEqualTo(2.0);
        await Assert.That(result.Values[1].TryGetError(out var error)).IsTrue();
        await Assert.That(error).IsEqualTo(Error.DivZero);
        await Assert.That(NumberAt(result, 2)).IsEqualTo(6.0);
    }

    [Test]
    public async Task BlankCellsParticipateAsBlank()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = Number(1);
        // A2 left blank
        sheet["A3"] = Number(3);
        var context = new EvaluationContext(workbook);

        // A1:A3 = 0 → blank equals 0 → [FALSE, TRUE, FALSE]
        var node = Equal(Range("A1", "A3", sheet), Number(0));

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out var result)).IsTrue();
        await Assert.That(result.Length).IsEqualTo(3);
        await Assert.That(BoolAt(result, 0)).IsFalse();
        await Assert.That(BoolAt(result, 1)).IsTrue(); // blank == 0
        await Assert.That(BoolAt(result, 2)).IsFalse();
    }

    [Test]
    public async Task NonEligibleNode_ReturnsFalse()
    {
        var (workbook, sheet) = ShowHideGrid();
        var context = new EvaluationContext(workbook);

        // A plain scalar function (SUM of scalars) is not an array → TryEvaluate false.
        await Assert.That(ArrayEvaluation.TryEvaluate(Sum(Number(1), Number(2)), context, out _)).IsFalse();

        // A bare scalar literal → not an array.
        await Assert.That(ArrayEvaluation.TryEvaluate(Number(5), context, out _)).IsFalse();
    }

    [Test]
    public async Task OpenRange_IsRefused()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = String("Show");
        sheet["A2"] = String("Hide");
        var context = new EvaluationContext(workbook);

        // A:A="Show" — the whole-column open range is out of the mini-CSE cost guard → refuse.
        // Column A is 1-based index 1; both column limits set, both row limits open → a whole column.
        var openColumn = OpenRangeReference.Create(1, 1, null, null, sheet.Name);
        var node = Equal(openColumn, String("Show"));

        await Assert.That(ArrayEvaluation.TryEvaluate(node, context, out _)).IsFalse();

        // Directly, too.
        await Assert.That(ArrayEvaluation.TryEvaluate(openColumn, context, out _)).IsFalse();
    }
}
