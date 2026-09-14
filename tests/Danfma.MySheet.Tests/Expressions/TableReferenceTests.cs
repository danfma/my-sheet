using System.Reflection;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Information;
using Danfma.MySheet.Expressions.Lookup;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;
using MemoryPack;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The <see cref="TableReference"/> node itself — resolution over the Phase 3 registry (Phase 4 T1), the
/// measured six-area geometry it resolves through <see cref="Table.GetRegion"/> (Phase 5 T1), the two
/// invariants of <see cref="TableReference.Evaluate"/>, the union tag and the wire. The parser has emitted
/// the node since the parse-arm commit, so a tree here is built by hand as a test-isolation choice — the
/// parse-level shapes live in StructuredReferenceTests.
/// </summary>
public class TableReferenceTests
{
    // Data!Tabela1 = A1:B4, header row 1 (Item / Valor), data rows 2..4 = a,10 / b,20 / c,30, no totals row.
    private static Workbook Fixture()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["A2"] = new StringValue("a");
        data["B2"] = new NumberValue(10);
        data["A3"] = new StringValue("b");
        data["B3"] = new NumberValue(20);
        data["A4"] = new StringValue("c");
        data["B4"] = new NumberValue(30);
        workbook.DefineTable("Tabela1", "Data", "A1:B4", ["Item", "Valor"]);
        return workbook;
    }

    // The oracle's own fixture, cell for cell: Data!Tabela1 = A1:C4, header Item / Valor / Qtd, data rows
    // a,10,1 / b,20,2 / c,30,3. With a totals row it is A1:C5 and Aspose's own ShowTotals row: A5 = "Total",
    // B5 empty (no total chosen for Valor), C5 = SUBTOTAL(109,[Qtd]) = 6, which is why the with-totals
    // numbers below are 72 / 14 and [[#Totals],[Valor]] sums to 0 over a real, resolvable B5:B5.
    private static Workbook Oracle(bool totals)
    {
        var workbook = new Workbook();
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

        if (totals)
        {
            data["A5"] = new StringValue("Total");
            data["C5"] = new NumberValue(6);
        }

        workbook.DefineTable(
            "Tabela1",
            "Data",
            totals ? "A1:C5" : "A1:C4",
            ["Item", "Valor", "Qtd"],
            hasHeaderRow: true,
            hasTotalsRow: totals
        );
        return workbook;
    }

    private static TableReference Node(
        string table = "Tabela1",
        string? column = "Valor",
        TableArea area = TableArea.Data
    ) => new(table, column, area);

    // === TryResolve: the one primitive ===================================================================

    [Test]
    public async Task Data_WithAColumn_ResolvesToTheColumnsDataRows()
    {
        var ok = Node().TryResolve(Fixture(), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    // T[#Data] and T[] (Aspose rewrites the latter to the bare table name): the whole data body.
    [Test]
    public async Task Data_WithNoColumn_ResolvesToTheWholeDataBody()
    {
        var ok = Node(column: null).TryResolve(Fixture(), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("A2", "B4", "Data"));
    }

    // Excel resolves table and column names case-insensitively (Table1[col] -> Table1[Col]).
    [Test]
    public async Task TableAndColumnNames_ResolveCaseInsensitively()
    {
        var ok = Node("TABELA1", "valor").TryResolve(Fixture(), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    // Unknown table -> #NAME?, the same name space as a defined name (NameReference).
    [Test]
    public async Task UnknownTable_IsName()
    {
        var ok = Node("NoSuch").TryResolve(Fixture(), out var range, out var error);

        await Assert.That(ok).IsFalse();
        await Assert.That(range).IsNull();
        await Assert.That(error).IsEqualTo(Error.Name);
    }

    // Unknown column -> #REF!, Excel's own repair (a deleted column's specifier becomes Table1[#REF!]).
    [Test]
    public async Task UnknownColumn_IsRef()
    {
        var ok = Node(column: "NoSuch").TryResolve(Fixture(), out var range, out var error);

        await Assert.That(ok).IsFalse();
        await Assert.That(range).IsNull();
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    // Sweep item 33, reopening Phase 5 ruling R1: the oracle treats a header-only table's data band as an
    // EMPTY reference (SUM 0, ROWS 0, ISREF TRUE — EmptyTableReferenceTests has the consumer matrix), so the
    // band resolves to a zero-row rectangle anchored under the header, over the column(s) it names. Before
    // the empty-reference representation this answered false / null / #REF!, the recorded divergence.
    [Test]
    [Arguments("Valor", 2, 2)]
    [Arguments(null, 1, 2)]
    public async Task HeaderOnlyTable_DataIsAnEmptyReference(string? column, int left, int right)
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Data");
        workbook.DefineTable("Vazia", "Data", "A1:B1", ["Item", "Valor"]);

        var ok = Node("Vazia", column).TryResolve(workbook, out var reference, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsEqualTo(new EmptyRangeReference("Data", 2, left, right));
    }

    // === The six areas, measured =========================================================================

    // Replaces Phase 4 T1's interim NonDataAreas_AreRef_UntilPhase5TryGetRegion, which asserted #REF! for
    // the five non-Data areas while only the [#Data] geometry existed. Every row below is the rectangle
    // Aspose.Cells 26.6.0 reports for the specifier (2026-09-11), read off ROW/ROWS/COLUMN/COLUMNS typed on
    // Main!H20 both PLAIN and array-entered — the two modes agreed everywhere. Oracle fixture: Data!Tabela1
    // = A1:C4, header Item/Valor/Qtd, data rows a,10,1 / b,20,2 / c,30,3; with a totals row it is A1:C5,
    // A5 = "Total", B5 empty, C5 = SUBTOTAL(109,[Qtd]) = 6.
    //
    // | specifier                    | no totals row | with a totals row |
    // | ---------------------------- | ------------- | ----------------- |
    // | [#All]                       | A1:C4         | A1:C5             |
    // | [#Data]                      | A2:C4         | A2:C4             |
    // | [#Headers]                   | A1:C1         | A1:C1             |
    // | [#Totals]                    | #REF!         | A5:C5             |
    // | [[#Headers],[#Data]]         | A1:C4         | A1:C4             |
    // | [[#Data],[#Totals]]          | A2:C4         | A2:C5             |
    // | [Valor] / [[#Data],[Valor]]  | B2:B4         | B2:B4             |
    // | [[#All],[Valor]]             | B1:B4         | B1:B5             |
    // | [[#Headers],[Valor]]         | B1:B1         | B1:B1             |
    // | [[#Totals],[Valor]]          | #REF!         | B5:B5             |
    // | [[#Headers],[#Data],[Valor]] | B1:B4         | B1:B4             |
    // | [[#Data],[#Totals],[Valor]]  | B2:B4         | B2:B5             |
    [Test]
    [Arguments(TableArea.All, null, false, "A1", "C4")]
    [Arguments(TableArea.All, null, true, "A1", "C5")]
    [Arguments(TableArea.Data, null, false, "A2", "C4")]
    [Arguments(TableArea.Data, null, true, "A2", "C4")]
    [Arguments(TableArea.Headers, null, false, "A1", "C1")]
    [Arguments(TableArea.Headers, null, true, "A1", "C1")]
    [Arguments(TableArea.Totals, null, true, "A5", "C5")]
    [Arguments(TableArea.HeadersAndData, null, false, "A1", "C4")]
    [Arguments(TableArea.HeadersAndData, null, true, "A1", "C4")]
    [Arguments(TableArea.DataAndTotals, null, false, "A2", "C4")]
    [Arguments(TableArea.DataAndTotals, null, true, "A2", "C5")]
    [Arguments(TableArea.Data, "Valor", false, "B2", "B4")]
    [Arguments(TableArea.Data, "Valor", true, "B2", "B4")]
    [Arguments(TableArea.All, "Valor", false, "B1", "B4")]
    [Arguments(TableArea.All, "Valor", true, "B1", "B5")]
    [Arguments(TableArea.Headers, "Valor", false, "B1", "B1")]
    [Arguments(TableArea.Totals, "Valor", true, "B5", "B5")]
    [Arguments(TableArea.HeadersAndData, "Valor", false, "B1", "B4")]
    [Arguments(TableArea.HeadersAndData, "Valor", true, "B1", "B4")]
    [Arguments(TableArea.DataAndTotals, "Valor", false, "B2", "B4")]
    [Arguments(TableArea.DataAndTotals, "Valor", true, "B2", "B5")]
    public async Task EveryArea_ResolvesToTheMeasuredRectangle(
        TableArea area,
        string? column,
        bool totals,
        string start,
        string end
    )
    {
        var ok = Node(column: column, area: area).TryResolve(Oracle(totals), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference(start, end, "Data"));
    }

    // The same rows as formula results, because a rectangle is only right if a consumer reads the cells the
    // oracle read: SUM / COUNTA / ROWS of each specifier, no totals row then with one. Both entry modes
    // agreed on every number here.
    [Test]
    [Arguments(TableArea.All, null, false, 66.0, 12.0, 4.0)]
    [Arguments(TableArea.All, null, true, 72.0, 14.0, 5.0)]
    [Arguments(TableArea.Data, null, false, 66.0, 9.0, 3.0)]
    [Arguments(TableArea.Data, null, true, 66.0, 9.0, 3.0)]
    [Arguments(TableArea.Headers, null, false, 0.0, 3.0, 1.0)]
    [Arguments(TableArea.Totals, null, true, 6.0, 2.0, 1.0)]
    [Arguments(TableArea.HeadersAndData, null, false, 66.0, 12.0, 4.0)]
    [Arguments(TableArea.HeadersAndData, null, true, 66.0, 12.0, 4.0)]
    [Arguments(TableArea.DataAndTotals, null, false, 66.0, 9.0, 3.0)]
    [Arguments(TableArea.DataAndTotals, null, true, 72.0, 11.0, 4.0)]
    [Arguments(TableArea.Data, "Valor", false, 60.0, 3.0, 3.0)]
    [Arguments(TableArea.All, "Valor", false, 60.0, 4.0, 4.0)]
    [Arguments(TableArea.All, "Valor", true, 60.0, 4.0, 5.0)]
    [Arguments(TableArea.Headers, "Valor", false, 0.0, 1.0, 1.0)]
    [Arguments(TableArea.Totals, "Valor", true, 0.0, 0.0, 1.0)]
    [Arguments(TableArea.DataAndTotals, "Valor", true, 60.0, 3.0, 4.0)]
    public async Task EveryArea_AnswersTheOraclesAggregates(
        TableArea area,
        string? column,
        bool totals,
        double sum,
        double countA,
        double rows
    )
    {
        var context = new EvaluationContext(Oracle(totals));
        var node = Node(column: column, area: area);

        await Assert.That(new Sum([node]).Evaluate(context).AsObject() as double?).IsEqualTo(sum);
        await Assert
            .That(new CountA([node]).Evaluate(context).AsObject() as double?)
            .IsEqualTo(countA);
        await Assert.That(new Rows([node]).Evaluate(context).AsObject() as double?).IsEqualTo(rows);
    }

    // Only the SINGLETONS error when the row they name is absent: [#Totals] and [[#Totals],[Valor]] over a
    // table with no totals row are #REF! through SUM/ROW/ROWS/COLUMN/COLUMNS, and ISREF is FALSE. Measured;
    // COUNTA is 1 in both, because COUNTA counts the error as one element (ruling R2's shape) — do not read
    // that 1 as an empty region.
    [Test]
    [Arguments(null)]
    [Arguments("Valor")]
    public async Task TotalsArea_WithNoTotalsRow_IsRefAndIsNotAReference(string? column)
    {
        var context = new EvaluationContext(Oracle(false));
        var node = Node(column: column, area: TableArea.Totals);

        await Assert.That(node.TryResolve(Oracle(false), out _, out var error)).IsFalse();
        await Assert.That(error).IsEqualTo(Error.Ref);
        await Assert
            .That(new Sum([node]).Evaluate(context).AsObject())
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(new Rows([node]).Evaluate(context).AsObject())
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(new IsRef([node]).Evaluate(context).AsObject() as bool?).IsFalse();
        await Assert
            .That(new CountA([node]).Evaluate(context).AsObject() as double?)
            .IsEqualTo(1.0);
    }

    // The pairs SHRINK where the singleton errors — the mistake Phase 4 ruling 2 names. Half one, measured
    // over the A1:C4 fixture with no totals row: SUM(T[[#Data],[#Totals]]) = 66, the data body, NOT #REF!.
    [Test]
    public async Task APairMissingItsNamedRow_IsTheDataBody_NotRef()
    {
        var workbook = Oracle(false);
        var context = new EvaluationContext(workbook);
        var pair = Node(column: null, area: TableArea.DataAndTotals);
        var data = Node(column: null, area: TableArea.Data);

        await Assert.That(pair.TryResolve(workbook, out var pairRange, out _)).IsTrue();
        await Assert.That(data.TryResolve(workbook, out var dataRange, out _)).IsTrue();
        await Assert.That(pairRange).IsEqualTo(dataRange);
        await Assert.That(new Sum([pair]).Evaluate(context).AsObject() as double?).IsEqualTo(66.0);
    }

    // Half two, the OTHER direction of the shrink rule and its own fixture: a header-LESS table. Aspose
    // cannot be asked for one directly — ShowHeaderRow = false consumes the first row — so the measured shape
    // is the table it leaves behind, two data rows at A2:C3 (b,20,2 / c,30,3). Over it,
    // SUM(T[[#Headers],[#Data]]) = 55, ROW 2, ROWS 2, i.e. the data body, while SUM(T[#Headers]) is #REF!:
    // the singleton errors, the pair that names it does not.
    [Test]
    public async Task AHeaderLessTable_HeadersIsRef_ButHeadersAndDataIsTheDataBody()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A2"] = new StringValue("b");
        data["B2"] = new NumberValue(20);
        data["C2"] = new NumberValue(2);
        data["A3"] = new StringValue("c");
        data["B3"] = new NumberValue(30);
        data["C3"] = new NumberValue(3);
        workbook.DefineTable(
            "Tabela1",
            "Data",
            "A2:C3",
            ["Item", "Valor", "Qtd"],
            hasHeaderRow: false,
            hasTotalsRow: false
        );
        var context = new EvaluationContext(workbook);

        var pair = Node(column: null, area: TableArea.HeadersAndData);
        await Assert.That(pair.TryResolve(workbook, out var range, out _)).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("A2", "C3", "Data"));
        await Assert.That(new Sum([pair]).Evaluate(context).AsObject() as double?).IsEqualTo(55.0);
        await Assert.That(new Rows([pair]).Evaluate(context).AsObject() as double?).IsEqualTo(2.0);

        var headers = Node(column: null, area: TableArea.Headers);
        await Assert.That(headers.TryResolve(workbook, out _, out var error)).IsFalse();
        await Assert.That(error).IsEqualTo(Error.Ref);
        await Assert
            .That(new Sum([headers]).Evaluate(context).AsObject())
            .IsEqualTo(ErrorValue.Reference);
    }

    // The second legal shape with an empty band, after the header-only table: a totals row, no header row and
    // no data (ref="A1:C1" totalsRowCount="1", which Validate accepts). Not measurable on the oracle — Aspose
    // cannot build a table with no data rows AND no header row — so this is MySheet's own answer from the one
    // geometry rule: the bands that need a data row are EMPTY, anchored at the first row the table has (they
    // were false / #REF! before sweep item 33), [#Headers] is absent (#REF!), and the rest are that row.
    [Test]
    public async Task ATotalsOnlyTable_TheDataBandsAreEmpty_AndTheRestIsTheTotalsRow()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Total");
        data["C1"] = new NumberValue(6);
        workbook.DefineTable(
            "SoTotais",
            "Data",
            "A1:C1",
            ["Item", "Valor", "Qtd"],
            hasHeaderRow: false,
            hasTotalsRow: true
        );

        foreach (var area in new[] { TableArea.Data, TableArea.HeadersAndData })
        {
            var node = new TableReference("SoTotais", null, area);
            await Assert.That(node.TryResolve(workbook, out var reference, out _)).IsTrue();
            await Assert.That(reference).IsEqualTo(new EmptyRangeReference("Data", 1, 1, 3));
        }

        await Assert
            .That(
                new TableReference("SoTotais", null, TableArea.Headers).TryResolve(
                    workbook,
                    out _,
                    out var headersError
                )
            )
            .IsFalse();
        await Assert.That(headersError).IsEqualTo(Error.Ref);

        foreach (var area in new[] { TableArea.All, TableArea.Totals, TableArea.DataAndTotals })
        {
            var node = new TableReference("SoTotais", null, area);
            await Assert.That(node.TryResolve(workbook, out var range, out _)).IsTrue();
            await Assert.That(range).IsEqualTo(new RangeReference("A1", "C1", "Data"));
        }
    }

    // Sweep item 33 through the node: on a header-only table the oracle answers an EMPTY reference for every
    // band that needs data (measured: SUM 0, COUNT 0, COUNTA 0, ROWS 0, COLUMNS 1, ISREF TRUE, SUBTOTAL(9) 0,
    // COUNTIF 0, AVERAGE #DIV/0!, INDEX(…,1,1) #REF!, ROW 2 PLAIN / 1 CSE), and both such bands now resolve to
    // the zero-row rectangle under the header (before: false / #REF!, ruling R1's recorded divergence). The
    // bands that do NOT need data still resolve exactly as the oracle has them: [#All], [#Headers] and
    // [[#Headers],[#Data]] are the header row (ROWS 1, COUNTA 3 on both).
    [Test]
    public async Task HeaderOnlyTable_TheDataBandsAreEmpty_AndTheHeaderBandsResolve()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Data");
        workbook.DefineTable("Vazia", "Data", "A1:C1", ["Item", "Valor", "Qtd"]);

        foreach (var area in new[] { TableArea.Data, TableArea.DataAndTotals })
        {
            var empty = new TableReference("Vazia", null, area);
            await Assert.That(empty.TryResolve(workbook, out var reference, out _)).IsTrue();
            await Assert.That(reference).IsEqualTo(new EmptyRangeReference("Data", 2, 1, 3));
        }

        foreach (var area in new[] { TableArea.All, TableArea.Headers, TableArea.HeadersAndData })
        {
            var header = new TableReference("Vazia", null, area);
            await Assert.That(header.TryResolve(workbook, out var range, out _)).IsTrue();
            await Assert.That(range).IsEqualTo(new RangeReference("A1", "C1", "Data"));
        }
    }

    // === Evaluate: the two invariants ====================================================================

    // (a) Evaluate returns the CONCRETE resolved range, never ComputedValue.Reference(this).
    [Test]
    public async Task Evaluate_ReturnsTheConcreteRange_NotItself()
    {
        var value = Node().Evaluate(new EvaluationContext(Fixture()));

        await Assert.That(value.TryGetReference(out var reference)).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    // The measured reason for (a): a probe returning Reference(this) made SUM(node) answer 0, because
    // EnumerateValues' catch-all `case Reference` yields the reference back as one non-numeric element.
    [Test]
    public async Task Sum_OverTheNode_ExpandsTheResolvedRange()
    {
        var sum = new Sum([Node()]);

        var value = sum.Evaluate(new EvaluationContext(Fixture())).AsObject();

        await Assert.That(value as double?).IsEqualTo(60.0);
    }

    // (b) Evaluate never answers #VALUE! the way RangeReference.Evaluate does, and (Phase 5 ruling R2) an
    // unresolvable table is an error VALUE that flows to the consumer, never a throw or a short-circuit.
    [Test]
    public async Task Evaluate_UnknownTable_IsTheNameErrorValue()
    {
        var context = new EvaluationContext(Fixture());

        var bare = Node("NoSuch").Evaluate(context).AsObject();
        var summed = new Sum([Node("NoSuch")]).Evaluate(context).AsObject();

        await Assert.That(bare).IsEqualTo(ErrorValue.Name);
        await Assert.That(summed).IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task Evaluate_UnknownColumn_IsTheRefErrorValue()
    {
        var value = Node(column: "NoSuch").Evaluate(new EvaluationContext(Fixture())).AsObject();

        await Assert.That(value).IsEqualTo(ErrorValue.Reference);
    }

    // === ISFORMULA / FORMULATEXT: item 17 ================================================================
    //
    // Oracle (Aspose.Cells 26.6.0, 2026-09-11, both entry modes): ISFORMULA(Tabela1[Valor]) is FALSE over
    // the fixture's plain-literal value cells and FORMULATEXT is #N/A (no formula to show, matching a
    // literal-range top-left). An unresolvable table (Tabela1[#Totals] with no totals row) is ISFORMULA
    // FALSE -- NOT the #VALUE! the two switches' `_ => (null, null)` default gives a non-reference argument
    // -- and FORMULATEXT #REF!, the node's own error VALUE, mirroring ReferenceGuard's TableReference arm
    // (R2): the failure IS a value, never a short-circuit to a different code.

    [Test]
    public async Task IsFormula_OverAValueCell_IsFalse()
    {
        var value = new IsFormula([Node()]).Evaluate(new EvaluationContext(Fixture())).AsObject();

        await Assert.That(value as bool?).IsFalse();
    }

    [Test]
    public async Task IsFormula_OverAnUnresolvableTable_IsFalse_NotAValueError()
    {
        var workbook = Oracle(false);
        var node = Node(column: null, area: TableArea.Totals);

        var value = new IsFormula([node]).Evaluate(new EvaluationContext(workbook)).AsObject();

        await Assert.That(value as bool?).IsFalse();
    }

    [Test]
    public async Task FormulaText_OverAValueCell_IsNotAvailable()
    {
        var value = new FormulaText([Node()]).Evaluate(new EvaluationContext(Fixture())).AsObject();

        await Assert.That(value).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task FormulaText_OverAnUnresolvableTable_IsTheNodesOwnError()
    {
        var workbook = Oracle(false);
        var node = Node(column: null, area: TableArea.Totals);

        var value = new FormulaText([node]).Evaluate(new EvaluationContext(workbook)).AsObject();

        await Assert.That(value).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task FormulaText_OverAnUnknownTable_IsTheNameError()
    {
        var value = new FormulaText([Node("NoSuch")])
            .Evaluate(new EvaluationContext(Fixture()))
            .AsObject();

        await Assert.That(value).IsEqualTo(ErrorValue.Name);
    }

    // Proves the arm actually resolves through the range's StartId/SheetName (not a hardcoded answer): B2
    // holds a real FORMULA here, so ISFORMULA flips to TRUE and FORMULATEXT round-trips its un-parsed text,
    // exactly like the AnchoredRangeReference arms this mirrors.
    [Test]
    public async Task IsFormula_AndFormulaText_ReadTheResolvedRangesTopLeftCell()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Item");
        data["B1"] = new StringValue("Valor");
        data["A2"] = new StringValue("a");
        data["B2"] = ExpressionParser.Parse("=5*2", data);
        data["A3"] = new StringValue("b");
        data["B3"] = new NumberValue(20);
        workbook.DefineTable("Tabela1", "Data", "A1:B3", ["Item", "Valor"]);
        var context = new EvaluationContext(workbook);
        var node = Node();

        var isFormula = new IsFormula([node]).Evaluate(context).AsObject();
        var formulaText = new FormulaText([node]).Evaluate(context).AsObject();

        await Assert.That(isFormula as bool?).IsTrue();
        await Assert.That(formulaText).IsEqualTo("=5*2");
    }

    // === TryResolveReference ============================================================================

    [Test]
    public async Task TryResolveReference_ResolvesToTheRange()
    {
        var ok = Node().TryResolveReference(new EvaluationContext(Fixture()), out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }

    [Test]
    public async Task TryResolveReference_Unresolvable_IsFalse()
    {
        var ok = Node("NoSuch")
            .TryResolveReference(new EvaluationContext(Fixture()), out var reference);

        await Assert.That(ok).IsFalse();
        await Assert.That(reference).IsNull();
    }

    // Deliberately NOT overridden: a table's target comes from REGISTERED state, invalidated by the
    // definitions version bump, not by per-pass volatility — overriding it would make DependencyExtractor
    // mark every structured-reference formula always-dirty and throw away its static RangeDep.
    [Test]
    public async Task IsVolatile_IsNotOverridden()
    {
        await Assert.That(Node().IsVolatile).IsFalse();
    }

    // === The enum and the wire ===========================================================================

    // Data = 0 so the overwhelmingly common T[Col] serializes the enum's default byte; the two pairs are
    // appended after the four singletons. Phase 4 ruling 1.
    [Test]
    public async Task TableArea_IsAByteEnum_WithDataAsDefault()
    {
        var members = Enum.GetValues<TableArea>()
            .Select(area => $"{area}={Convert.ToByte(area)}")
            .ToArray();

        await Assert.That(Enum.GetUnderlyingType(typeof(TableArea))).IsEqualTo(typeof(byte));
        await Assert
            .That(members)
            .IsEquivalentTo([
                "Data=0",
                "All=1",
                "Headers=2",
                "Totals=3",
                "HeadersAndData=4",
                "DataAndTotals=5",
            ]);
        await Assert.That(default(TableArea).ToString()).IsEqualTo("Data");
    }

    [Test]
    public async Task SerializationRoundTrip_PreservesEveryMember()
    {
        var node = new TableReference("Tabela1", "Sales Amount", TableArea.HeadersAndData);
        var bytes = MemoryPackSerializer.Serialize<Expression>(node);
        var back = MemoryPackSerializer.Deserialize<Expression>(bytes);

        await Assert.That(back).IsEqualTo(node);
    }

    [Test]
    public async Task SerializationRoundTrip_PreservesANullColumn()
    {
        var node = new TableReference("Tabela1", null, TableArea.All);
        var bytes = MemoryPackSerializer.Serialize<Expression>(node);
        var back = (TableReference)MemoryPackSerializer.Deserialize<Expression>(bytes)!;

        await Assert.That(back.ColumnName).IsNull();
        await Assert.That(back.Area).IsEqualTo(TableArea.All);
    }

    // Union tags are APPEND-ONLY. 322 is Aggregate (Phase 2), 323-326 are Phase 7's producers, so the
    // node took 327 — the number the design wrote (322) was already live. A duplicate tag fails at
    // MemoryPack type initialization, not at compile time, which is why the tag is pinned by value.
    [Test]
    public async Task UnionTag_Is327()
    {
        var tag = typeof(Expression)
            .GetCustomAttributes<MemoryPackUnionAttribute>()
            .Single(attribute => attribute.Type == typeof(TableReference))
            .Tag;

        await Assert.That(tag).IsEqualTo((ushort)327);
    }

    // Sweep item 33's empty reference is a runtime VALUE a structured reference resolves to — never parsed,
    // never stored in a cell tree — so it is deliberately not a union member and the wire does not move.
    [Test]
    public async Task EmptyRangeReference_IsNotAUnionMember()
    {
        var members = typeof(Expression)
            .GetCustomAttributes<MemoryPackUnionAttribute>()
            .Select(attribute => attribute.Type);

        await Assert.That(members.Contains(typeof(EmptyRangeReference))).IsFalse();
    }
}
