using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// <see cref="NamedReferences.CaptureValue"/> is the single rule behind "a range node bound to a name / chosen
/// by CHOOSE / passed through unary + stays a range". These tests cover the arms the LET tests do not: CHOOSE
/// over a union, and the shared-formula <see cref="AnchoredRangeReference"/> arm, which must resolve to the
/// rectangle of THE SLAVE being evaluated, not the master's.
/// </summary>
public class CaptureValueTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");
        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Number(3);
        sheet["A4"] = Number(4);

        return (workbook, sheet);
    }

    // Phase 5 items 20/23's canonical fixture: Data!Tabela1 = A1:C4, header Item/Valor/Qtd, rows a,10,1 /
    // b,20,2 / c,30,3. The parser cannot spell Tabela1[Valor] yet (Phase 4 T5 owns Parser.cs), so every tree
    // below is built by hand, exactly as TableReferenceTests.cs does.
    private static Workbook TableWorkbook()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["C1"] = new StringValue("Qtd");
        data["A2"] = new StringValue("a");
        data["B2"] = Number(10);
        data["C2"] = Number(1);
        data["A3"] = new StringValue("b");
        data["B3"] = Number(20);
        data["C3"] = Number(2);
        data["A4"] = new StringValue("c");
        data["B4"] = Number(30);
        data["C4"] = Number(3);
        workbook.DefineTable("Tabela1", "Data", "A1:C4", ["Item", "Valor", "Qtd"]);
        return workbook;
    }

    // Tabela1[Valor] (B2:B4 = 10, 20, 30) — the column every test below reads.
    private static TableReference Valor => new("Tabela1", "Valor", TableArea.Data);

    [Test]
    public async Task Choose_Union_StaysARange()
    {
        var (workbook, sheet) = Grid();

        var value = ExpressionParser
            .Parse("=SUM(CHOOSE(1,(A1:A1,A3:A4)))", sheet)
            .Evaluate(workbook);

        await Assert.That(value.AsDouble()).IsEqualTo(8.0);
    }

    [Test]
    [Arguments("SUM(CHOOSE(1,A1:A2))")]
    [Arguments("LET(r,A1:A2,SUM(r))")]
    [Arguments("SUM(+A1:A2)")]
    public async Task AnchoredRange_InSharedFormula_ResolvesPerSlave(string body)
    {
        // Master at row 1 reads A1:A2 (=3); the slave one row down must read A2:A3 (=5), i.e. the anchored
        // range is shifted by the slave's delta before being captured — not frozen at the master's rectangle.
        var (workbook, sheet) = Grid();
        var master = ExpressionParser.ParseAnchoredMasterBody(
            ExpressionParser.TokenizeFormulaBody(body),
            sheet
        );

        var atMaster = new SharedFormulaSlave(master, 0, 0).Evaluate(workbook);
        var oneRowDown = new SharedFormulaSlave(master, 1, 0).Evaluate(workbook);

        await Assert.That(atMaster.AsDouble()).IsEqualTo(3.0);
        await Assert.That(oneRowDown.AsDouble()).IsEqualTo(5.0);
    }

    // === Phase 5 item 20: a structured reference through the same four binding sites ====================
    //
    // TableReference has no arm of its own in CaptureValue's switch, so all four pass through the shared
    // `_ => expression.Evaluate(context)` fallback — the same route a bare RangeReference/OpenRangeReference
    // takes. That is enough because TableReference.Evaluate already yields the reference VALUE for a
    // resolved table (never a scalar), so the fallback captures a RANGE, not a number. These four are the
    // guard against ever changing that Evaluate to RangeReference.Evaluate's #VALUE! convention: if it did,
    // every one of these would collapse to one #VALUE!-turned-error cell instead of streaming the column.

    [Test]
    public async Task StructuredReference_BoundToALetName_StaysARange()
    {
        // =LET(r,Tabela1[Valor],SUM(r)) -> 60
        var r = new NameReference("r");
        var formula = new Let([r, Valor, new Sum([r])]);

        var value = formula.Evaluate(new EvaluationContext(TableWorkbook()));

        await Assert.That(value.AsDouble()).IsEqualTo(60.0);
    }

    [Test]
    public async Task StructuredReference_ThroughChoose_StaysARange()
    {
        // =SUM(CHOOSE(1,Tabela1[Valor])) -> 60
        var formula = new Sum([new Choose([Number(1), Valor])]);

        var value = formula.Evaluate(new EvaluationContext(TableWorkbook()));

        await Assert.That(value.AsDouble()).IsEqualTo(60.0);
    }

    [Test]
    public async Task StructuredReference_ThroughUnaryPlus_StaysARange()
    {
        // =SUM(+Tabela1[Valor]) -> 60. Unary + is Excel's reference-preserving no-op; CaptureValue has no
        // UnaryOperation arm either, so this falls through the SAME default to UnaryOperation.Evaluate,
        // which for `+` returns its operand's own value unchanged.
        var formula = new Sum([new UnaryOperation(UnaryOperator.Plus, Valor)]);

        var value = formula.Evaluate(new EvaluationContext(TableWorkbook()));

        await Assert.That(value.AsDouble()).IsEqualTo(60.0);
    }

    [Test]
    public async Task StructuredReference_AsADefinedNameDefinition_StaysARange()
    {
        // A defined name whose definition is a structured reference: =SUM(ProdName) -> 60.
        // NamedReferences.EvaluateDefinition (the scalar reading of a name) routes through
        // ArrayBindings.Capture, which for a bare-reference definition (a TableReference IS one) takes
        // CaptureValue's path too — the same fallback, the same guard.
        var workbook = TableWorkbook();
        workbook.DefineName("ProdName", Valor);
        var formula = new Sum([new NameReference("ProdName")]);

        var value = formula.Evaluate(new EvaluationContext(workbook));

        await Assert.That(value.AsDouble()).IsEqualTo(60.0);
    }
}
