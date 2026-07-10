using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Expressions;

public class DynamicRangeTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return (workbook, sheet);
    }

    [Test]
    public async Task Resolves_IndexEndpoint_ToSpanningRange()
    {
        var (workbook, sheet) = Grid();
        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        // INDEX(A1:A3,2) -> A2 ; span A2:A3
        var range = new DynamicRange(
            ExpressionParser.Parse("=INDEX(A1:A3,2)", sheet),
            new CellReference("A3", "Sheet1")
        );

        var ok = range.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        var rr = (RangeReference)reference!;
        await Assert.That(rr.StartId).IsEqualTo("A2");
        await Assert.That(rr.EndId).IsEqualTo("A3");
    }

    [Test]
    public async Task SerializationRoundTrip_PreservesEndpoints()
    {
        var range = new DynamicRange(
            new CellReference("A1", "Sheet1"),
            new CellReference("A3", "Sheet1")
        );
        var bytes = MemoryPackSerializer.Serialize<Expression>(range);
        var back = (DynamicRange)MemoryPackSerializer.Deserialize<Expression>(bytes)!;

        await Assert.That(((CellReference)back.Start).Id).IsEqualTo("A1");
        await Assert.That(((CellReference)back.End).Id).IsEqualTo("A3");
    }
}
