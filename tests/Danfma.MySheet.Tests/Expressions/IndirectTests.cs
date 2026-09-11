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
        ExpressionParser
            .Parse(formula, sheet)
            .Evaluate(new EvaluationContext(workbook, sheet.Name))
            .AsObject();

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
        await Assert
            .That(Calc(wb, sheet, "=INDIRECT(\"not a ref\")"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Indirect_R1C1Style_IsRefError()
    {
        // MySheet supports A1 style only; a1 = FALSE (R1C1) is #REF!.
        var (wb, sheet) = Grid(("A1", 5));
        await Assert
            .That(Calc(wb, sheet, "=INDIRECT(\"A1\", FALSE)"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Indirect_SerializationRoundTrip()
    {
        var node = new Indirect([new Danfma.MySheet.Expressions.StringValue("A1")]);
        var bytes = MemoryPackSerializer.Serialize<Expression>(node);
        var back = MemoryPackSerializer.Deserialize<Expression>(bytes);
        await Assert.That(back).IsTypeOf<Indirect>();
    }

    // === Phase 5 item 23: INDIRECT over a structured reference ==========================================
    //
    // The brief's claim, now TRUE: Phase 4's parser arms merged, so ParseFormulaBody(refText, sheet)
    // (Indirect.cs) produces a TableReference and TryResolveReference resolves it exactly as a defined
    // name does. These were canaries pinned at #REF! while the grammar could not be reached (the interim
    // head threw ParseException from ParseFormulaBody, which Indirect's own catch turned into false);
    // after the Phase 4 merge they answered 60 immediately, measured before flipping — the oracle's
    // number, stored identically (Aspose.Cells 26.6.0, both entry modes agree on this shape).

    [Test]
    public async Task StructuredReference_ThroughIndirect_Resolves()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Item");
        data["B1"] = new Danfma.MySheet.Expressions.StringValue("Valor");
        data["A2"] = new Danfma.MySheet.Expressions.StringValue("a");
        data["B2"] = new NumberValue(10);
        data["A3"] = new Danfma.MySheet.Expressions.StringValue("b");
        data["B3"] = new NumberValue(20);
        data["A4"] = new Danfma.MySheet.Expressions.StringValue("c");
        data["B4"] = new NumberValue(30);
        workbook.DefineTable("Tabela1", "Data", "A1:B4", ["Item", "Valor"]);

        // Lives in a real cell, read via GetCellValue — Indirect.TryResolveReference needs a current sheet
        // (context.SheetName), which a bare Expression.Evaluate(workbook) would not supply.
        data["H1"] = ExpressionParser.Parse("=SUM(INDIRECT(\"Tabela1[Valor]\"))", data);

        var value = workbook.GetCellValue("Data", "H1");

        await Assert.That(value.AsDouble()).IsEqualTo(60.0);
    }

    [Test]
    public async Task StructuredReference_ThroughIndirect_WithARuntimeColumnName_Resolves()
    {
        // The corpus shape: =SUM(INDIRECT("Tabela1["&SUBSTITUTE(D1,"'","''")&"]")) with D1 holding a column
        // name containing an apostrophe — the round trip of the lexer's escape table (item 1) against
        // SUBSTITUTE's output: D1 -> O''Col inside the brackets, decoded back to O'Col at resolution.
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Item");
        data["B1"] = new Danfma.MySheet.Expressions.StringValue("O'Col");
        data["A2"] = new Danfma.MySheet.Expressions.StringValue("a");
        data["B2"] = new NumberValue(10);
        data["A3"] = new Danfma.MySheet.Expressions.StringValue("b");
        data["B3"] = new NumberValue(20);
        data["A4"] = new Danfma.MySheet.Expressions.StringValue("c");
        data["B4"] = new NumberValue(30);
        workbook.DefineTable("Tabela1", "Data", "A1:B4", ["Item", "O'Col"]);

        data["D1"] = new Danfma.MySheet.Expressions.StringValue("O'Col"); // the raw column name, apostrophe included
        data["H1"] = ExpressionParser.Parse(
            "=SUM(INDIRECT(\"Tabela1[\"&SUBSTITUTE(D1,\"'\",\"''\")&\"]\"))",
            data
        );

        var value = workbook.GetCellValue("Data", "H1");

        await Assert.That(value.AsDouble()).IsEqualTo(60.0);
    }
}
