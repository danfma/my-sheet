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
    public async Task Offset_MultiCell_ResolvesToRangeReference()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=OFFSET(A1,0,0,2,2)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsTypeOf<RangeReference>();
        var range = (RangeReference)reference!;
        await Assert.That(range.StartId).IsEqualTo("A1");
        await Assert.That(range.EndId).IsEqualTo("B2");
    }

    [Test]
    public async Task Offset_DivZeroInRowsArgument_EvaluatesToSpecificError()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=OFFSET(A1,1/0,0)", sheet);

        var result = expr.Evaluate(new EvaluationContext(workbook));

        await Assert.That(result.TryGetError(out var error)).IsTrue();
        await Assert.That(error).IsEqualTo(Error.DivZero);
    }

    [Test]
    public async Task Offset_NonNumericRowsArgument_EvaluatesToSpecificError()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=OFFSET(A1,\"x\",0)", sheet);

        var result = expr.Evaluate(new EvaluationContext(workbook));

        await Assert.That(result.TryGetError(out var error)).IsTrue();
        await Assert.That(error).IsEqualTo(Error.Value);
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

    [Test]
    public async Task Choose_ResolvesChosenArgumentToReference()
    {
        var (workbook, sheet) = Grid();
        var expr = ExpressionParser.Parse("=CHOOSE(2,A1,B5)", sheet);

        var ok = expr.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(((CellReference)reference!).Id).IsEqualTo("B5");
    }

    // === Structured references ===========================================================================

    // One fixture for the bounds matrix below: Data!Tabela1 = A1:B5 with a header row AND a totals row, so
    // header 1, data 2..4, totals 5, columns Item / Valor. Every expected rectangle is the one Aspose.Cells
    // 26.6.0 reports for the specifier (measured 2026-09-11 on the same shape: ROW/ROWS/COLUMN/COLUMNS typed
    // PLAIN and array-entered, which agreed). The nodes are built by hand exactly as TableReferenceTests
    // builds them — an isolation choice now that the parser has its bracket arm (Phase 4).
    private static Workbook TableFixture(bool totals = true, string valorColumn = "Valor")
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new Danfma.MySheet.Expressions.StringValue("Item");
        data["B1"] = new Danfma.MySheet.Expressions.StringValue(valorColumn);
        data["A2"] = new Danfma.MySheet.Expressions.StringValue("a");
        data["B2"] = new NumberValue(10);
        data["A3"] = new Danfma.MySheet.Expressions.StringValue("b");
        data["B3"] = new NumberValue(20);
        data["A4"] = new Danfma.MySheet.Expressions.StringValue("c");
        data["B4"] = new NumberValue(30);

        if (totals)
        {
            data["A5"] = new Danfma.MySheet.Expressions.StringValue("Total");
            // Deliberately NOT 60, the sum of the data rows a real totals row would show: a rectangle that
            // wrongly slid onto or grew into the totals row must change every number below, and a fixture
            // value that coincides with the right answer would hide it.
            data["B5"] = new NumberValue(999);
        }

        workbook.DefineTable(
            "Tabela1",
            "Data",
            totals ? "A1:B5" : "A1:B4",
            ["Item", valorColumn],
            hasHeaderRow: true,
            hasTotalsRow: totals
        );
        return workbook;
    }

    private static async Task AssertResolves(
        Workbook workbook,
        Expression node,
        string start,
        string end
    )
    {
        var ok = node.TryResolveReference(new EvaluationContext(workbook), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference(start, end, "Data"));
    }

    private static async Task AssertError(Workbook workbook, Expression node, Error expected)
    {
        var context = new EvaluationContext(workbook);

        await Assert.That(node.TryResolveReference(context, out var reference)).IsFalse();
        await Assert.That(reference).IsNull();
        await Assert.That(node.Evaluate(context).TryGetError(out var error)).IsTrue();
        await Assert.That(error).IsEqualTo(expected);
    }

    // Tabela1[Valor]: the column's DATA cells, header and totals excluded.
    [Test]
    public async Task Table_ColumnSpecifier_ResolvesToTheColumnsDataRows() =>
        await AssertResolves(
            TableFixture(),
            new TableReference("Tabela1", "Valor", TableArea.Data),
            "B2",
            "B4"
        );

    // Tabela1[#Data] (and Tabela1[], which Excel stores as the bare table name): the whole data body.
    [Test]
    public async Task Table_DataItem_ResolvesToAllDataRows() =>
        await AssertResolves(
            TableFixture(),
            new TableReference("Tabela1", null, TableArea.Data),
            "A2",
            "B4"
        );

    // Tabela1[#All]: "the entire table, including column headers, data, and totals".
    [Test]
    public async Task Table_AllItem_IncludesHeaderAndTotals() =>
        await AssertResolves(
            TableFixture(),
            new TableReference("Tabela1", null, TableArea.All),
            "A1",
            "B5"
        );

    [Test]
    public async Task Table_HeadersItem_IsTheHeaderRowOnly() =>
        await AssertResolves(
            TableFixture(),
            new TableReference("Tabela1", null, TableArea.Headers),
            "A1",
            "B1"
        );

    [Test]
    public async Task Table_TotalsItem_IsTheTotalsRowOnly() =>
        await AssertResolves(
            TableFixture(),
            new TableReference("Tabela1", null, TableArea.Totals),
            "A5",
            "B5"
        );

    // Tabela1[Valor] and Tabela1[[#Data],[Valor]] are ONE node, not two that agree: both are
    // (Data, "Valor"), and Data is the default byte on the wire. Measured, the oracle agrees that the two
    // written forms are the same reference — B2:B4 either way (ROW 2, ROWS 3, COLUMN 2, COLUMNS 1), SUM 60,
    // COUNTA 3 — so the numbers below are the ones BOTH forms must answer. Aspose stores
    // [[#Data],[Valor]] back verbatim while MySheet will render both as [Valor] (a writer divergence the
    // phase notes; measured: Tabela1[[Valor]] and Tabela1[] are normalized by Aspose too, to
    // Tabela1[Valor] and the bare Tabela1). The two parser arms that must converge on this node are Phase 4
    // T5's and are pinned there — what is pinned here is the geometry they will both land on.
    [Test]
    public async Task Table_DataAndColumnComposite_EqualsThePlainColumnForm()
    {
        var workbook = TableFixture();
        var context = new EvaluationContext(workbook);
        var node = new TableReference("Tabela1", "Valor", TableArea.Data);

        await Assert.That(node.Area).IsEqualTo(default(TableArea));
        await AssertResolves(workbook, node, "B2", "B4");
        await Assert.That(new Sum([node]).Evaluate(context).AsObject() as double?).IsEqualTo(60.0);
        await Assert
            .That(new CountA([node]).Evaluate(context).AsObject() as double?)
            .IsEqualTo(3.0);
    }

    // An unknown table is #NAME?, because Excel resolves a table name in the same name space as a defined
    // name (NameReference answers #NAME? for an unknown one).
    [Test]
    public async Task Table_UnknownTable_IsName() =>
        await AssertError(
            TableFixture(),
            new TableReference("NoSuch", "Valor", TableArea.Data),
            Error.Name
        );

    // An unknown column is #REF!, Excel's own repair marker (a deleted column's specifier becomes
    // Table1[#REF!]). Unmeasurable on the oracle directly — Aspose rejects the formula at entry with
    // "Invalid table column: NoSuchCol" — so the code follows the measured analogue for an unknown name.
    [Test]
    public async Task Table_UnknownColumn_IsRef() =>
        await AssertError(
            TableFixture(),
            new TableReference("Tabela1", "NoSuch", TableArea.Data),
            Error.Ref
        );

    // Measured, both entry modes: SUM(Tabela1[#Totals]) over a table with no totals row is #REF! and ISREF
    // is FALSE. Only the singletons error this way; the pairs shrink (TableReferenceTests pins that).
    [Test]
    public async Task Table_TotalsItem_WithNoTotalsRow_IsRef() =>
        await AssertError(
            TableFixture(totals: false),
            new TableReference("Tabela1", null, TableArea.Totals),
            Error.Ref
        );

    // The lexer stores the DECODED payload — Tabela1[O''Brien] carries the column name O'Brien — because the
    // lookup compares against the raw tableColumn/@name of the xlsx, which holds the apostrophe unescaped.
    // The second half of the assertion is the coordination point: a node still carrying the written form
    // O''Brien matches nothing and is #REF!, so a lexer that forgets to decode fails here rather than
    // silently resolving. Same escape the corpus leans on in
    // INDIRECT("Overflow2["&SUBSTITUTE(D$1,"'","''")&"]").
    [Test]
    public async Task Table_ColumnWithEscapedApostrophe_MatchesTheDecodedName()
    {
        var workbook = TableFixture(valorColumn: "O'Brien");

        await AssertResolves(
            workbook,
            new TableReference("Tabela1", "O'Brien", TableArea.Data),
            "B2",
            "B4"
        );
        await AssertError(
            workbook,
            new TableReference("Tabela1", "O''Brien", TableArea.Data),
            Error.Ref
        );
    }
}
