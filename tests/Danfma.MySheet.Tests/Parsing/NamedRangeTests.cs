using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Workbook-level defined names: a <see cref="NameReference"/> resolves against
/// <see cref="Workbook.DefinedNames"/> (after the LET scope), and the functions that require a syntactic
/// reference node (VLOOKUP/HLOOKUP table, INDEX, OFFSET base, ROWS, COLUMNS, AREAS, ISREF) unwrap a name to
/// the range/cell it stands for. Semantics mirror Excel; the LET-shadowing and cycle cases are called out.
/// </summary>
public class NamedRangeTests
{
    // Data!A1:A3 = 10/20/30, Data!B1:C3 a 3x2 grid, plus a scalar Data!Z1 = 0.1. Formulas parse on Main.
    private static (Workbook Workbook, Sheet Main) Book()
    {
        var workbook = new Workbook();

        var data = workbook.Sheets.Add("Data");
        data["A1"] = new NumberValue(10);
        data["A2"] = new NumberValue(20);
        data["A3"] = new NumberValue(30);
        data["B1"] = new NumberValue(1);
        data["C1"] = new StringValue("a");
        data["B2"] = new NumberValue(2);
        data["C2"] = new StringValue("b");
        data["B3"] = new NumberValue(3);
        data["C3"] = new StringValue("c");
        data["Z1"] = new NumberValue(0.1);

        return (workbook, workbook.Sheets.Add("Main"));
    }

    private static object? Eval(Workbook workbook, Sheet sheet, string formula) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

    [Test]
    public async Task Name_ToRange_FeedsAggregation()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");

        // SUM(name) expands the named range, exactly like SUM over the literal range.
        await Assert.That(Eval(workbook, main, "=SUM(Vendas)") as double?).IsEqualTo(60.0);
    }

    [Test]
    public async Task Name_CaseInsensitive()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");

        // Names are case-insensitive, like Excel: vendas == VENDAS == Vendas.
        await Assert.That(Eval(workbook, main, "=SUM(vendas)") as double?).IsEqualTo(60.0);
        await Assert.That(Eval(workbook, main, "=SUM(VENDAS)") as double?).IsEqualTo(60.0);
    }

    [Test]
    public async Task Name_ToTable_VLookup()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Tabela", "Data!B1:C3");

        // VLOOKUP resolves the named table's first column and returns column 2 of the matching row.
        await Assert
            .That(Eval(workbook, main, "=VLOOKUP(2,Tabela,2,FALSE)") as string)
            .IsEqualTo("b");
    }

    [Test]
    public async Task Name_ToTable_HLookup()
    {
        var (workbook, main) = Book();
        // A 2x3 horizontal table: keys in the first row, results below.
        var data = workbook["Data"];
        data["E1"] = new NumberValue(1);
        data["F1"] = new NumberValue(2);
        data["G1"] = new NumberValue(3);
        data["E2"] = new StringValue("x");
        data["F2"] = new StringValue("y");
        data["G2"] = new StringValue("z");
        workbook.DefineName("Horizontal", "Data!E1:G2");

        await Assert
            .That(Eval(workbook, main, "=HLOOKUP(2,Horizontal,2,FALSE)") as string)
            .IsEqualTo("y");
    }

    [Test]
    public async Task Name_ToTable_Index()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Tabela", "Data!B1:C3");

        // INDEX(Tabela, 1, 2) is row 1, column 2 of B1:C3 -> C1 = "a".
        await Assert.That(Eval(workbook, main, "=INDEX(Tabela,1,2)") as string).IsEqualTo("a");
    }

    [Test]
    public async Task Name_ToRange_Rows()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");

        await Assert.That(Eval(workbook, main, "=ROWS(Vendas)") as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Name_ToTable_Columns()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Tabela", "Data!B1:C3");

        await Assert.That(Eval(workbook, main, "=COLUMNS(Tabela)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Name_ToCell_Offset()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Origin", "Data!A1");

        // OFFSET(Origin, 2, 0) from Data!A1 is Data!A3 = 30.
        await Assert.That(Eval(workbook, main, "=OFFSET(Origin,2,0)") as double?).IsEqualTo(30.0);
    }

    [Test]
    public async Task Name_ToRange_Areas()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");
        workbook.DefineName("Blocos", "(Data!A1:A3,Data!B1:B3)");

        await Assert.That(Eval(workbook, main, "=AREAS(Vendas)") as double?).IsEqualTo(1.0);
        await Assert.That(Eval(workbook, main, "=AREAS(Blocos)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Name_ToRange_IsRef_And_Constant_IsNotRef()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");
        workbook.DefineName("Taxa", new NumberValue(0.1));

        // A name that stands for a reference is a reference; a name for a constant is not (per Excel).
        await Assert.That((bool)Eval(workbook, main, "=ISREF(Vendas)")!).IsTrue();
        await Assert.That((bool)Eval(workbook, main, "=ISREF(Taxa)")!).IsFalse();
    }

    [Test]
    public async Task Name_ToSingleCell_Scalar()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Taxa", "Data!Z1");

        // A name for a single cell evaluates to that cell's scalar value.
        await Assert.That(Eval(workbook, main, "=Taxa*100") as double?).IsEqualTo(10.0);
    }

    [Test]
    public async Task Name_ToConstant_AnyExpression()
    {
        var (workbook, main) = Book();
        // Names can point to ANY expression, not only references.
        workbook.DefineName("Taxa", new NumberValue(0.1));

        await Assert.That(Eval(workbook, main, "=Taxa*100") as double?).IsEqualTo(10.0);
    }

    [Test]
    public async Task Let_ShadowsDefinedName()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");

        // The LET binding shadows the workbook name: Vendas is 5 here, not the range, so the result is 6.
        await Assert
            .That(Eval(workbook, main, "=LET(Vendas,5,Vendas+1)") as double?)
            .IsEqualTo(6.0);
    }

    [Test]
    public async Task NameCycle_YieldsRefError_NoOverflow()
    {
        var (workbook, main) = Book();
        workbook.DefineName("A", "=B");
        workbook.DefineName("B", "=A");

        // A <-> B is a name cycle: it must degrade to #REF!, not overflow the stack.
        await Assert.That(Eval(workbook, main, "=A")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task UndefinedName_IsNameError()
    {
        var (workbook, main) = Book();

        // Regression: an unknown bare name is still #NAME?.
        await Assert.That(Eval(workbook, main, "=Missing")).IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task DefineName_TransitiveName_ToRange()
    {
        var (workbook, main) = Book();
        workbook.DefineName("Vendas", "Data!A1:A3");
        workbook.DefineName("Alias", "=Vendas");

        // A name that points at another name for a range still expands in an aggregation.
        await Assert.That(Eval(workbook, main, "=SUM(Alias)") as double?).IsEqualTo(60.0);
    }

    [Test]
    public async Task DefineName_EmptyName_Throws()
    {
        var workbook = new Workbook();

        await Assert.That(() => workbook.DefineName("", "Data!A1")).Throws<ArgumentException>();
        await Assert.That(() => workbook.DefineName("   ", "Data!A1")).Throws<ArgumentException>();
    }

    [Test]
    public async Task DefineName_CellShapedName_Throws()
    {
        var workbook = new Workbook();

        // "A1" collides with a cell reference; Excel reserves that shape.
        await Assert
            .That(() => workbook.DefineName("A1", new NumberValue(1)))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task DefineName_UnqualifiedReference_Throws()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Data");

        // Names are workbook-level: an unqualified range has no sheet, so it is rejected.
        await Assert
            .That(() => workbook.DefineName("Vendas", "A1:A10"))
            .Throws<ArgumentException>();
        await Assert
            .That(() => workbook.DefineName("Total", "=SUM(A1:A10)"))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task NameReference_RoundTripsThroughFormulaWriter()
    {
        var (_, main) = Book();

        // The parser already produces NameReference for a bare name, and the FormulaWriter prints it back
        // verbatim, so a formula using a name round-trips unchanged (no parser change was needed).
        const string formula = "SUM(Vendas)";
        var parsed = ExpressionParser.Parse("=" + formula, main);

        await Assert.That(parsed.ToFormula(main.Name)).IsEqualTo(formula);
    }

    [Test]
    public async Task DefineName_QualifiedReference_Succeeds()
    {
        var (workbook, main) = Book();

        // A fully-qualified reference (with or without '$' and leading '=') is accepted.
        workbook.DefineName("Vendas", "Data!$A$1:$A$3");
        workbook.DefineName("Origin", "=Data!$A$1");

        await Assert.That(Eval(workbook, main, "=SUM(Vendas)") as double?).IsEqualTo(60.0);
        await Assert.That(Eval(workbook, main, "=Origin*2") as double?).IsEqualTo(20.0);
    }
}
