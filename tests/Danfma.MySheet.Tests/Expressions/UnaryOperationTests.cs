using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class UnaryOperationTests
{
    private static readonly Workbook Workbook = new();

    [Test]
    public async Task Negate_NegatesNumber()
    {
        await Assert
            .That(Negate(Number(2)).Evaluate(Workbook).AsObject() as double?)
            .IsEqualTo(-2.0);
    }

    [Test]
    public async Task Negate_Stacked_CancelsOut()
    {
        // -(-2) == 2
        await Assert
            .That(Negate(Negate(Number(2))).Evaluate(Workbook).AsObject() as double?)
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task Plus_KeepsNumber()
    {
        await Assert.That(Plus(Number(3)).Evaluate(Workbook).AsObject() as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Negate_PropagatesError()
    {
        await Assert
            .That(Negate(Divide(Number(1), Number(0))).Evaluate(Workbook).AsObject())
            .IsEqualTo(ErrorValue.DivByZero);
    }

    // --- Unary '+' is Excel's legacy (Lotus) no-op: it returns the operand unchanged, TYPE included.
    // Only a blank becomes 0. Regression for issue #8 (defect 1): `=+A1` on a text cell was #VALUE!.

    [Test]
    public async Task Plus_KeepsText()
    {
        await Assert
            .That(Plus(String("some text")).Evaluate(Workbook).AsObject())
            .IsEqualTo("some text");
    }

    [Test]
    public async Task Plus_KeepsNumericLookingText_AsText()
    {
        // "42" stays TEXT: the no-op must not silently coerce a numeric-looking string.
        var value = Plus(String("42")).Evaluate(Workbook);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Text);
        await Assert.That(value.AsString()).IsEqualTo("42");
    }

    [Test]
    public async Task Plus_KeepsBooleanType()
    {
        // Aspose/Excel keep TRUE as a boolean; a numeric coercion would turn it into 1.
        var value = Plus(BooleanValue.True).Evaluate(Workbook);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(value.AsBoolean()).IsTrue();
    }

    [Test]
    public async Task Plus_BlankIsZero()
    {
        var value = Plus(BlankValue.Instance).Evaluate(Workbook);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.AsDouble()).IsEqualTo(0.0);
    }

    [Test]
    public async Task Plus_PropagatesError()
    {
        await Assert
            .That(Plus(Divide(Number(1), Number(0))).Evaluate(Workbook).AsObject())
            .IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Plus_OnTextCell_ReturnsTheText()
    {
        // The issue's exact repro shape: `=+A1` where A1 holds text.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");
        sheet["A1"] = String("some text");

        var value = ExpressionParser.Parse("=+A1", sheet).Evaluate(workbook);

        await Assert.That(value.AsString()).IsEqualTo("some text");
    }

    [Test]
    public async Task Plus_PassesReferenceThrough_ToRangeConsumers()
    {
        // A reference-typed operand (OFFSET's multi-cell result) is passed through untouched, so
        // SUM(+OFFSET(...)) still sees a range — the no-op must not collapse it into #VALUE!.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");
        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Number(3);

        var value = ExpressionParser.Parse("=SUM(+OFFSET(A1,0,0,3,1))", sheet).Evaluate(workbook);

        await Assert.That(value.AsDouble()).IsEqualTo(6.0);
    }

    // --- A range NODE under '+' is still a range for every consumer (external review of PR #9 caught that
    // only reference VALUES from OFFSET/INDEX passed through; a syntactic A1:A3 became #VALUE!).

    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");
        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Number(3);
        sheet["B1"] = Number(10);
        sheet["C1"] = Number(20);
        sheet["D1"] = String("42");

        return (workbook, sheet);
    }

    [Test]
    [Arguments("=SUM(+A1:A3)", 6.0)] // value-path consumer, bounded range
    [Arguments("=SUM(+A:A)", 6.0)] // whole column
    [Arguments("=SUM(+(A1:A2,A3:A3))", 6.0)] // union
    [Arguments("=MATCH(20,+A1:C1,0)", 3.0)] // lookup array
    [Arguments("=ROWS(+A1:A3)", 3.0)] // syntactic consumer (TryResolveReference)
    [Arguments("=INDEX(+A1:A3,2)", 2.0)]
    [Arguments("=SUM(-(+A1:A3))", -6.0)] // Phase 11c item 11/2: '+' transparent to the probe, '-' lifts through it
    public async Task Plus_OnRangeNode_IsStillARange(string formula, double expected)
    {
        var (workbook, sheet) = Grid();

        var value = ExpressionParser.Parse(formula, sheet).Evaluate(workbook);

        await Assert.That(value.AsDouble()).IsEqualTo(expected);
    }

    [Test]
    public async Task Plus_OnRangeNode_UsedAsScalar_IsValueError()
    {
        // Exactly like `=A1:A3+1`: the no-op does not make a range usable where a scalar is required.
        var (workbook, sheet) = Grid();

        var value = ExpressionParser.Parse("=(+A1:A3)+1", sheet).Evaluate(workbook);

        await Assert.That(value.AsObject()).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Plus_OnRangeToMissingSheet_IsRefError()
    {
        // Structural #REF! must survive the no-op: SUM(+Ghost!A1:A3) == SUM(Ghost!A1:A3), not an empty 0.
        var (workbook, sheet) = Grid();

        var value = ExpressionParser.Parse("=SUM(+Ghost!A1:A3)", sheet).Evaluate(workbook);

        await Assert.That(value.AsObject()).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Plus_OnNumericLookingTextCell_StaysText()
    {
        // The issue measured cells, not literals: D1 holds the TEXT "42".
        var (workbook, sheet) = Grid();

        var value = ExpressionParser.Parse("=+D1", sheet).Evaluate(workbook);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Text);
        await Assert.That(value.AsString()).IsEqualTo("42");
    }

    // --- Unary '-' and postfix '%' keep coercing to a number (text -> #VALUE!), unlike '+'.

    [Test]
    public async Task Negate_OnRangeNode_LiftsElementwiseInAConsumedPosition()
    {
        // Before Phase 8 (elementwise lifting) this pinned #VALUE!: the Negate was evaluated once over the
        // range node, which has no scalar value. Inside an array-consuming argument the mini-CSE now lifts
        // '-' over each cell (Excel: -1,-2,-3, summed), while unary '+' keeps its reference-preserving path.
        var (workbook, sheet) = Grid();

        var value = ExpressionParser.Parse("=SUM(-A1:A3)", sheet).Evaluate(workbook);

        await Assert.That(value.AsDouble()).IsEqualTo(-6.0);
    }

    [Test]
    public async Task Negate_OnText_IsValueError()
    {
        await Assert
            .That(Negate(String("some text")).Evaluate(Workbook).AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Percent_OnText_IsValueError()
    {
        var percent = new UnaryOperation(UnaryOperator.Percent, String("some text"));

        await Assert.That(percent.Evaluate(Workbook).AsObject()).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Negate_OnBoolean_CoercesToNumber()
    {
        await Assert
            .That(Negate(BooleanValue.True).Evaluate(Workbook).AsObject() as double?)
            .IsEqualTo(-1.0);
    }
}
