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

    // --- Unary '-' and postfix '%' keep coercing to a number (text -> #VALUE!), unlike '+'.

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
