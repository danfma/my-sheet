using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Syntactic consumers over whole-column / whole-row references: VLOOKUP/HLOOKUP table, INDEX, OFFSET
/// base, AREAS, ISREF resolve through the POPULATED bounding box; ROWS/COLUMNS use the populated extent on
/// an OPEN axis and the exact structural count on a BOUNDED axis (a documented divergence from Excel's
/// fixed grid).
/// </summary>
public class WholeColumnConsumerTests
{
    private static (Workbook Workbook, Sheet Sheet) Sheet(
        params (string Id, Expression Value)[] cells
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return (workbook, sheet);
    }

    private static object? Eval(string formula, Sheet sheet, Workbook workbook) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

    // --- ROWS / COLUMNS ---

    [Test]
    public async Task Rows_WholeColumn_IsPopulatedRowExtent()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        await Assert.That(Eval("=ROWS(A:A)", sheet, workbook) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Rows_SparseColumn_SpansMinToMax()
    {
        var (workbook, sheet) = Sheet(("A5", Number(1)), ("A10", Number(2)));

        await Assert.That(Eval("=ROWS(A:A)", sheet, workbook) as double?).IsEqualTo(6.0);
    }

    [Test]
    public async Task Rows_EmptyColumn_IsZero()
    {
        var (workbook, sheet) = Sheet(("B1", Number(1)));

        await Assert.That(Eval("=ROWS(A:A)", sheet, workbook) as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task Columns_MultiColumn_IsStructuralAndExact()
    {
        // COLUMNS(A:C) = 3 even though only A is populated: a bounded axis is structural.
        var (workbook, sheet) = Sheet(("A1", Number(1)));

        await Assert.That(Eval("=COLUMNS(A:C)", sheet, workbook) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Columns_WholeColumn_IsOne()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        await Assert.That(Eval("=COLUMNS(A:A)", sheet, workbook) as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task Rows_WholeRow_IsStructural()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)));

        await Assert.That(Eval("=ROWS(1:5)", sheet, workbook) as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task Columns_WholeRow_IsPopulatedColumnExtent()
    {
        // 1:5 has an OPEN column axis; COLUMNS is the populated column extent (B..D = 3).
        var (workbook, sheet) = Sheet(("B1", Number(1)), ("D3", Number(2)));

        await Assert.That(Eval("=COLUMNS(1:5)", sheet, workbook) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Columns_EmptyWholeRow_IsZero()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)));

        await Assert.That(Eval("=COLUMNS(2:2)", sheet, workbook) as double?).IsEqualTo(0.0);
    }

    // --- VLOOKUP / HLOOKUP / INDEX / OFFSET ---

    [Test]
    public async Task VLookup_WholeColumnTable()
    {
        var (workbook, sheet) = Sheet(
            ("A1", Number(1)),
            ("B1", String("one")),
            ("A2", Number(2)),
            ("B2", String("two"))
        );

        await Assert
            .That(Eval("=VLOOKUP(2,A:B,2,FALSE)", sheet, workbook) as string)
            .IsEqualTo("two");
    }

    [Test]
    public async Task Index_WholeColumn_ByPopulatedPosition()
    {
        var (workbook, sheet) = Sheet(("A1", Number(10)), ("A2", Number(20)), ("A3", Number(30)));

        await Assert.That(Eval("=INDEX(A:A,3)", sheet, workbook) as double?).IsEqualTo(30.0);
    }

    [Test]
    public async Task Offset_WholeColumnBase()
    {
        // OFFSET base A:A resolves to the populated box (A1:A2); offset (row 1, col 0) from its start = A2.
        var (workbook, sheet) = Sheet(("A1", Number(10)), ("A2", Number(20)));

        await Assert.That(Eval("=OFFSET(A:A,1,0)", sheet, workbook) as double?).IsEqualTo(20.0);
    }

    // --- AREAS / ISREF ---

    [Test]
    public async Task Areas_WholeColumn_IsOne()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)));

        await Assert.That(Eval("=AREAS(A:A)", sheet, workbook) as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task IsRef_WholeColumn_IsTrue()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)));

        await Assert.That(Eval("=ISREF(A:A)", sheet, workbook) as bool?).IsTrue();
    }

    [Test]
    public async Task IsRef_EmptyWholeColumn_IsTrue()
    {
        var (workbook, sheet) = Sheet(("B1", Number(1)));

        await Assert.That(Eval("=ISREF(A:A)", sheet, workbook) as bool?).IsTrue();
    }
}
