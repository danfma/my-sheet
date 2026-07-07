using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Expressions;

public class IndirectTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid(params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        foreach (var (id, value) in cells)
        {
            sheet[id] = new NumberValue(value);
        }

        return (workbook, sheet);
    }

    private static object? Calc(Workbook workbook, Sheet sheet, string formula) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(new EvaluationContext(workbook, sheet.Name)).AsObject();

    [Test]
    public async Task Indirect_A1Style_ResolvesCell()
    {
        var (wb, sheet) = Grid(("A1", 5));
        await Assert.That(Calc(wb, sheet, "=INDIRECT(\"A1\")") as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task Indirect_QualifiedSheet_Resolves()
    {
        var (wb, sheet) = Grid();
        var other = wb.Sheets.Add("Sheet2");
        other["B2"] = new NumberValue(7);
        await Assert.That(Calc(wb, sheet, "=INDIRECT(\"Sheet2!B2\")") as double?).IsEqualTo(7.0);
    }

    [Test]
    public async Task Indirect_Range_SumsThroughIt()
    {
        var (wb, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));
        await Assert.That(Calc(wb, sheet, "=SUM(INDIRECT(\"A1:A3\"))") as double?).IsEqualTo(6.0);
    }

    [Test]
    public async Task Indirect_AsRangeEndpoint_Spans()
    {
        // INDIRECT("A1") is a reference endpoint of the ':' operator: INDIRECT("A1"):A3 == A1:A3.
        var (wb, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));
        await Assert.That(Calc(wb, sheet, "=SUM(INDIRECT(\"A1\"):A3)") as double?).IsEqualTo(6.0);
    }

    [Test]
    public async Task Indirect_InvalidText_IsRefError()
    {
        var (wb, sheet) = Grid();
        await Assert.That(Calc(wb, sheet, "=INDIRECT(\"not a ref\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Indirect_R1C1Style_IsRefError()
    {
        // MySheet supports A1 style only; a1 = FALSE (R1C1) is #REF!.
        var (wb, sheet) = Grid(("A1", 5));
        await Assert.That(Calc(wb, sheet, "=INDIRECT(\"A1\", FALSE)")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Indirect_SerializationRoundTrip()
    {
        var node = new Indirect([new Danfma.MySheet.Expressions.StringValue("A1")]);
        var bytes = MemoryPackSerializer.Serialize<Expression>(node);
        var back = MemoryPackSerializer.Deserialize<Expression>(bytes);
        await Assert.That(back).IsTypeOf<Indirect>();
    }
}
