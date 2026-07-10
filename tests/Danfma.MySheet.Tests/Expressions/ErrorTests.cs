using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Tests.Expressions;

public class ErrorTests
{
    [Test]
    public async Task Display_MatchesExcelCodes()
    {
        await Assert.That(Error.Null.Display).IsEqualTo("#NULL!");
        await Assert.That(Error.DivZero.Display).IsEqualTo("#DIV/0!");
        await Assert.That(Error.Value.Display).IsEqualTo("#VALUE!");
        await Assert.That(Error.Ref.Display).IsEqualTo("#REF!");
        await Assert.That(Error.Name.Display).IsEqualTo("#NAME?");
        await Assert.That(Error.Num.Display).IsEqualTo("#NUM!");
        await Assert.That(Error.NA.Display).IsEqualTo("#N/A");
    }

    [Test]
    public async Task ToString_PrintsDisplay()
    {
        await Assert.That(Error.DivZero.ToString()).IsEqualTo("#DIV/0!");
        await Assert.That($"{Error.Value}").IsEqualTo("#VALUE!");
    }

    [Test]
    public async Task Equality_ByCode()
    {
        await Assert.That(Error.FromDisplay("#VALUE!") == Error.Value).IsTrue();
        await Assert.That(Error.Value != Error.Num).IsTrue();
        await Assert.That(Error.Value.Equals(Error.Value)).IsTrue();
        await Assert.That(Error.Value.Equals(Error.Ref)).IsFalse();
    }

    [Test]
    public async Task RoundTrips_ThroughErrorValueNode()
    {
        // Error -> ErrorValue (nó de AST) -> Error, preservando a identidade.
        await Assert
            .That(Error.FromDisplay(Error.DivZero.ToErrorValue().ErrorCode))
            .IsEqualTo(Error.DivZero);
        await Assert.That(Error.FromDisplay(Error.NA.ToErrorValue().ErrorCode)).IsEqualTo(Error.NA);
    }
}
