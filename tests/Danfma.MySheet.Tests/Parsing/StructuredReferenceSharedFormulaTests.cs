using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Mathematics;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Phase 4 item 11: a structured reference on the shared-formula FAST path.
/// <see cref="AnchoredFormulaSupport.IsFullyAnchored"/> ACCEPTS a <see cref="TableReference"/> for the reason it
/// accepts a <see cref="NameReference"/>: the node resolves by table name against the workbook-scoped registry and
/// carries no position component, so one master tree is correct for every slave and a table formula shared down a
/// column stays off the per-slave token re-parse (<c>ExpressionParser.ParseSharedFormulaBody</c>).
/// <para>Every tree here is built by hand, exactly as <c>TableReferenceTests</c> does — the parser emits the node
/// since Phase 4 T5, and <c>StructuredReferenceTests</c> pins that its anchored-master and shifted-slave parses
/// produce the identical one, so nothing is lost by keeping this file's trees parser-free. The per-row behaviour is
/// read through
/// <see cref="Workbook.GetCellValue(string,string)"/> — the only path that crosses the cell boundary's implicit
/// intersection (see <c>CellBoundaryIntersectionTests</c>). The verdict tests and the per-row tests are
/// deliberately separate: the verdict is what <c>WorksheetStreamLoader.TryBuildAnchoredMaster</c> consults, and
/// the per-row tests are the measured reason the verdict is "accept" rather than "reject".</para>
/// </summary>
public class StructuredReferenceSharedFormulaTests
{
    // Data!Tabela1 = A1:B4, header row 1 (Item / Valor), data rows 2..4 = a,10 / b,20 / c,30, no totals row.
    // The formula cells live in column D of the same sheet, rows 2..5: three INSIDE the table's data rows and
    // one (row 5) OUTSIDE them.
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

    private static TableReference Valor() => new("Tabela1", "Valor", TableArea.Data);

    // Stores `master` at D2 and one slave per row below it (delta = row - 2), the way the loader's ExpandSlave
    // shapes a group whose master is anchored-safe, then reads every cell back through the cell boundary.
    private static object?[] SharedDownColumnD(Expression master, int lastRow)
    {
        var workbook = Fixture();
        var sheet = workbook["Data"];
        sheet["D2"] = master;

        for (var row = 3; row <= lastRow; row++)
        {
            sheet[$"D{row}"] = new SharedFormulaSlave(master, row - 2, 0);
        }

        return Enumerable
            .Range(2, lastRow - 1)
            .Select(row => workbook.GetCellValue("Data", $"D{row}").AsObject())
            .ToArray();
    }

    // === The verdict AnchoredFormulaSupport gives ========================================================

    [Test]
    public async Task IsFullyAnchored_AcceptsABareTableReference()
    {
        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(Valor())).IsTrue();
    }

    // No area and no column makes the node position-dependent: [@Col], the only position-dependent structured
    // form, is out of scope and has no TableArea member.
    [Test]
    [Arguments("Valor", TableArea.Data)]
    [Arguments(null, TableArea.Data)]
    [Arguments(null, TableArea.All)]
    [Arguments(null, TableArea.Headers)]
    [Arguments(null, TableArea.Totals)]
    [Arguments("Valor", TableArea.HeadersAndData)]
    [Arguments("Valor", TableArea.DataAndTotals)]
    public async Task IsFullyAnchored_AcceptsEveryArea_WithOrWithoutAColumn(
        string? column,
        TableArea area
    )
    {
        var node = new TableReference("Tabela1", column, area);

        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(node)).IsTrue();
    }

    // The commonest real shapes: SUM(Tabela1[Valor]) and Tabela1[Valor]*2, shared down a column.
    [Test]
    public async Task IsFullyAnchored_AcceptsTheNode_InsideAFunctionAndAnOperator()
    {
        var summed = new Sum([Valor()]);
        var doubled = new BinaryOperation(BinaryOperator.Multiply, Valor(), new NumberValue(2));

        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(summed)).IsTrue();
        await Assert.That(AnchoredFormulaSupport.IsFullyAnchored(doubled)).IsTrue();
    }

    // === The per-row behaviour that makes "accept" correct ===============================================

    // Measured on Aspose.Cells 26.6.0 (2026-09-11) over this fixture, PLAIN entry: a bare Tabela1[Valor] in
    // D2 / D3 / D4 answers 10 / 20 / 30 and in D5 (below the table) #VALUE!; a SetSharedFormula group E2:E6
    // over the same text answers the same 10 / 20 / 30 / #VALUE! / #VALUE!, before and after a save + reload
    // that keeps it shared. (Array-entered, the same cells answer 10 in every row — the 1x1 array's top-left —
    // which is a different entry mode and is not compared here.) One shared master tree reproduces the
    // plain/shared column exactly, because the slave's delta shifts nothing in the node and the cell
    // boundary's implicit intersection runs per cell from the slave's own row.
    [Test]
    public async Task SharedGroup_OverABareTableColumn_IntersectsAtEachSlavesOwnRow()
    {
        var values = SharedDownColumnD(Valor(), lastRow: 5);

        await Assert.That(values[0]).IsEqualTo(10.0);
        await Assert.That(values[1]).IsEqualTo(20.0);
        await Assert.That(values[2]).IsEqualTo(30.0);
        await Assert.That(values[3]).IsEqualTo(ErrorValue.NotValue);
    }

    // A discriminating pin for "the delta shifts nothing": were the slave's rows applied to the resolved
    // rectangle (B2:B4 becoming B5:B7 three rows down), the row-5 slave would answer 0 (blank B5 coerced),
    // not #VALUE!, and SUM would answer 0 instead of 60 in that slave. Measured on the same oracle run, a
    // SetSharedFormula group G2:G5 over SUM(Tabela1[Valor]) answers 60 in all four cells.
    [Test]
    public async Task SharedGroup_OverSumOfATableColumn_IsTheWholeColumnInEverySlave()
    {
        var values = SharedDownColumnD(new Sum([Valor()]), lastRow: 5);

        await Assert.That(values).IsEquivalentTo(new object?[] { 60.0, 60.0, 60.0, 60.0 });
    }

    // The node-level statement of the same fact, without a cell: a slave with any delta resolves the table
    // column to the SAME rectangle the bare node does.
    [Test]
    [Arguments(0, 0)]
    [Arguments(3, 0)]
    [Arguments(3, 2)]
    public async Task Slave_ResolvesTheSameRectangle_WhateverItsDelta(int deltaRow, int deltaColumn)
    {
        var context = new EvaluationContext(Fixture());
        var slave = new SharedFormulaSlave(Valor(), deltaRow, deltaColumn);

        var ok = slave.TryResolveReference(context, out var reference);

        await Assert.That(ok).IsTrue();
        await Assert.That(reference).IsEqualTo(new RangeReference("B2", "B4", "Data"));
    }
}
