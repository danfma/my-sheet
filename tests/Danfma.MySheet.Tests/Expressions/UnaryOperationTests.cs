using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class UnaryOperationTests
{
    private static readonly Workbook Workbook = new();

    [Test]
    public async Task Negate_NegatesNumber()
    {
        await Assert.That(Negate(Number(2)).Compute(Workbook) as double?).IsEqualTo(-2.0);
    }

    [Test]
    public async Task Negate_Stacked_CancelsOut()
    {
        // -(-2) == 2
        await Assert.That(Negate(Negate(Number(2))).Compute(Workbook) as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Plus_KeepsNumber()
    {
        await Assert.That(Plus(Number(3)).Compute(Workbook) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Negate_PropagatesError()
    {
        await Assert
            .That(Negate(Divide(Number(1), Number(0))).Compute(Workbook))
            .IsEqualTo(ErrorValue.DivByZero);
    }
}
