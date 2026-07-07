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

    [Test]
    public async Task IndexIntoConcreteRange_ResolvesToTargetCell()
    {
        var (workbook, sheet) = Grid();
        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        var expr = ExpressionParser.Parse("=INDEX(A1:A3,2)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsTypeOf<CellReference>();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A2");
    }

    [Test]
    public async Task IndexIntoReversedRange_ResolvesToNormalizedTargetCell()
    {
        var (workbook, sheet) = Grid();
        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        // A3:A1 is a reversed range: StartId="A3", EndId="A1". Row 1 of the normalized range is A1, not A3.
        var expr = ExpressionParser.Parse("=INDEX(A3:A1,1,1)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsTypeOf<CellReference>();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A1");
    }

    [Test]
    public async Task Offset_ResolvesToTargetCell()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=OFFSET(A1,1,0)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("A2");
    }

    [Test]
    public async Task IndexIntoComputedArray_DoesNotResolve()
    {
        var (workbook, sheet) = Grid();
        // INDEX(ROW($A:$A), 2) indexes a computed vector, not cells: no address, must not resolve.
        var expr = ExpressionParser.Parse("=INDEX(ROW($A:$A),2)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsFalse();
        await Assert.That(reference).IsNull();
    }
}
