using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class TryResolveReferenceTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return (workbook, sheet);
    }

    [Test]
    public async Task CellReference_ResolvesToItself()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=A2", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsTypeOf<CellReference>();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A2");
    }

    [Test]
    public async Task NumberValue_DoesNotResolve()
    {
        var (workbook, _) = Grid();
        var expr = new NumberValue(3);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsFalse();
        await Assert.That(reference).IsNull();
    }
}
