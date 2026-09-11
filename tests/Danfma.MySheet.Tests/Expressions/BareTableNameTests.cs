using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 5 T5, ruling R3 mechanism (a): a BARE table name — <c>=Tabela1</c>, the spelling Excel itself
/// stores after normalizing <c>Tabela1[]</c> — must resolve through the NAME path AT EVALUATION TIME. The
/// parser holds a sheet and no workbook, so it cannot know that <c>Tabela1</c> is a table; what it CAN do
/// (and what this phase pins) is stop classifying the token as a CELL: <c>ParseIdentifier</c> routes an
/// identifier that is cell-shaped but outside Excel's grid (<c>Tabela1</c>, <c>XFE1</c>, <c>A1048577</c> —
/// exactly the set <c>Table.ValidateName</c> and <c>NamedReferences.IsValidName</c> reserve) to a
/// <see cref="NameReference"/>, and the name path (<c>NamedReferences.TryResolveRaw</c> /
/// <see cref="NameReference.Evaluate"/>) resolves a name registered in <c>Workbook.Tables</c> to
/// <c>TableReference(name, null, TableArea.Data)</c> — the CONCRETE data-body rectangle, mirroring
/// DynamicRange.
/// <para>
/// Oracle for every number: <b>Aspose.Cells 26.6.0, 2026-09-11</b>, over Data!Tabela1 = A1:C4 (header
/// Item/Valor/Qtd, data rows a,10,1 / b,20,2 / c,30,3), formulas typed on Main!H20 — far from the table's
/// rows, so the PLAIN cell-boundary rule has nothing to intersect and the bare 2-D name answers #VALUE!
/// (that row lives in CellBoundaryIntersectionTests, the only read path that crosses EvaluateCell). The
/// rows here evaluate the parsed expression directly, which is the ARRAY-ENTERED rule inside a function
/// argument — the mode the mini-CSE implements; the one row where the modes split is named.
/// </para>
/// </summary>
public class BareTableNameTests
{
    private static Workbook Fixture()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Main");
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["C1"] = new StringValue("Qtd");
        data["A2"] = new StringValue("a");
        data["B2"] = new NumberValue(10);
        data["C2"] = new NumberValue(1);
        data["A3"] = new StringValue("b");
        data["B3"] = new NumberValue(20);
        data["C3"] = new NumberValue(2);
        data["A4"] = new StringValue("c");
        data["B4"] = new NumberValue(30);
        data["C4"] = new NumberValue(3);
        workbook.DefineTable("Tabela1", "Data", "A1:C4", ["Item", "Valor", "Qtd"]);
        return workbook;
    }

    private static object? Calc(string formula)
    {
        var workbook = Fixture();
        var main = workbook.Sheets["Main"];

        return ExpressionParser
            .Parse(formula, main)
            .Evaluate(new EvaluationContext(workbook, "Main"))
            .AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // === The classification, and the resolution it feeds =================================================

    // The whole ruling in one assertion: the token is no longer a cell. A CellReference("TABELA1") reads an
    // always-blank cell and every row below answered a silent 0 before the routing changed.
    [Test]
    public async Task ABareTableStyleName_Parses_ToANameReference_NotACell()
    {
        var parsed = ExpressionParser.Parse("=Tabela1", new Sheet { Name = "Main" });

        await Assert.That(parsed).IsEqualTo(new NameReference("Tabela1"));
    }

    // The name path hands back the table's CONCRETE data body — TableReference(name, null, Data)'s own
    // rectangle, A2:C4 on the table's sheet — never a wrapper around the name node.
    [Test]
    public async Task ABareTableName_ResolvesToTheTablesConcreteDataBody()
    {
        var workbook = Fixture();
        var main = workbook.Sheets["Main"];
        var parsed = ExpressionParser.Parse("=Tabela1", main);

        var ok = parsed.TryResolveReference(
            new EvaluationContext(workbook, "Main"),
            out var reference
        );

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference("A2", "C4", "Data"));
    }

    // Table names resolve case-insensitively (the registry is OrdinalIgnoreCase; the oracle agrees,
    // Table1[table] -> Table1[Col] measured).
    [Test]
    public async Task ABareTableName_ResolvesCaseInsensitively()
    {
        await Assert.That(Num(Calc("=SUM(tabela1)"))).IsEqualTo(66.0);
    }

    // === The oracle's consumer rows ======================================================================
    //
    // | formula                        | PLAIN | CSE |
    // | ------------------------------ | ----- | --- |
    // | SUM(Tabela1)                   | 66    | 66  |
    // | ROWS(Tabela1)                  | 3     | 3   |
    // | COLUMNS(Tabela1)               | 3     | 3   |
    // | COUNTA(Tabela1)                | 9     | 9   |
    // | AREAS(Tabela1)                 | 1     | 1   |
    // | INDEX(Tabela1,2,2)             | 20    | 20  |
    // | COUNTIF(Tabela1,">15")         | 2     | 2   |
    // | LET(t,Tabela1,SUM(t))          | 66    | 66  |
    // | SUM(INDIRECT("Tabela1"))       | 66    | 66  |
    // | COUNT((Tabela1<>"")*1)         | 0     | 9   |  <- the one split; CSE asserted (see below)
    // | =Tabela1 bare                  | #VALUE! | "a" | <- cell-boundary row, CellBoundaryIntersectionTests
    [Test]
    [Arguments("=SUM(Tabela1)", 66.0)]
    [Arguments("=ROWS(Tabela1)", 3.0)]
    [Arguments("=COLUMNS(Tabela1)", 3.0)]
    [Arguments("=COUNTA(Tabela1)", 9.0)]
    [Arguments("=AREAS(Tabela1)", 1.0)]
    [Arguments("=INDEX(Tabela1,2,2)", 20.0)]
    [Arguments(@"=COUNTIF(Tabela1,"">15"")", 2.0)]
    [Arguments("=LET(t,Tabela1,SUM(t))", 66.0)]
    [Arguments(@"=SUM(INDIRECT(""Tabela1""))", 66.0)]
    public async Task TheConsumerRows_AnswerTheOracle(string formula, double expected)
    {
        await Assert.That(Num(Calc(formula))).IsEqualTo(expected);
    }

    // The CSE half of the one split row. Direct evaluation is the array-entered rule inside the argument
    // (the mini-CSE's Probe reaches the name through ResolveNameShape, which resolves it to the table's
    // rectangle exactly as it resolves a range-valued defined name), so 9 is the answer on this path. The
    // PLAIN column's 0 belongs to the cell boundary and is not this path's number.
    [Test]
    public async Task Count_OfComparisonOverABareTableName_IsElementWise()
    {
        await Assert.That(Num(Calc(@"=COUNT((Tabela1<>"""")*1)"))).IsEqualTo(9.0);
    }

    [Test]
    public async Task IsRef_OfABareTableName_IsTrue()
    {
        await Assert.That(Calc("=ISREF(Tabela1)") as bool?).IsTrue();
    }

    // The control that keeps the Tables arm honest: a name that is NEITHER a defined name nor a table
    // still answers #NAME? — the arm must not swallow the unknown-name case.
    [Test]
    public async Task AnUnknownName_StillAnswersName()
    {
        await Assert.That(Calc("=SUM(NoSuch)")).IsEqualTo(ErrorValue.Name);
    }
}
