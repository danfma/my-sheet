using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class BinaryOperationTests
{
    private static readonly Workbook Workbook = new();

    [Test]
    public async Task Add_ComputesSum()
    {
        var expr = Add(Number(1), Number(2));

        await Assert.That(expr.Evaluate(Workbook).AsObject() as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Subtract_IsLeftAssociativeWhenNested()
    {
        // (2 - 3) - 4 == -5
        var expr = Subtract(Subtract(Number(2), Number(3)), Number(4));

        await Assert.That(expr.Evaluate(Workbook).AsObject() as double?).IsEqualTo(-5.0);
    }

    [Test]
    public async Task Power_Computes()
    {
        var expr = Power(Number(2), Number(3));

        await Assert.That(expr.Evaluate(Workbook).AsObject() as double?).IsEqualTo(8.0);
    }

    [Test]
    public async Task Divide_ByZero_ReturnsDivByZeroError()
    {
        var expr = Divide(Number(1), Number(0));

        await Assert.That(expr.Evaluate(Workbook).AsObject()).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task GreaterThan_ReturnsBoolean()
    {
        var expr = GreaterThan(Number(2), Number(1));

        await Assert.That(expr.Evaluate(Workbook).AsObject() as bool?).IsTrue();
    }

    [Test]
    public async Task Arithmetic_PropagatesErrorFromOperand()
    {
        // 1 + (1 / 0) propagates #DIV/0!
        var expr = Add(Number(1), Divide(Number(1), Number(0)));

        await Assert.That(expr.Evaluate(Workbook).AsObject()).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task NonNumericString_CoercesToValueError()
    {
        var expr = Add(new Danfma.MySheet.Expressions.StringValue("abc"), Number(1));

        await Assert.That(expr.Evaluate(Workbook).AsObject()).IsEqualTo(ErrorValue.NotValue);
    }
}
