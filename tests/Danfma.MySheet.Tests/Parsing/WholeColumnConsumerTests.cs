using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Syntactic consumers over whole-column / whole-row references. VLOOKUP/HLOOKUP's table and AREAS/ISREF
/// resolve through the POPULATED bounding box. INDEX and OFFSET's base address ABSOLUTE positions instead
/// — column A / row 1 when a side is open, sweep item 37 follow-up ruling (a)
/// (<see cref="OpenRangeReference.AbsoluteRow"/>/<see cref="OpenRangeReference.AbsoluteColumn"/>): the two
/// readings coincide whenever data starts at the sheet's own row 1 / column A, which is why
/// <see cref="Index_WholeColumn_ByAbsolutePosition"/> pins BOTH a fixture where they agree and one where
/// they do not. ROWS/COLUMNS use the populated extent on an OPEN axis and the exact structural count on a
/// BOUNDED axis (a documented divergence from Excel's fixed grid) — unrelated to, and unaffected by, the
/// ruling above (ROWS/COLUMNS report an EXTENT, not a position to translate).
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
    public async Task Index_WholeColumn_ByAbsolutePosition()
    {
        // Data starting at row 1: absolute row 3 and "the 3rd populated row" are the SAME cell (A3), so
        // this fixture alone cannot tell the two conventions apart.
        var (workbook, sheet) = Sheet(("A1", Number(10)), ("A2", Number(20)), ("A3", Number(30)));

        await Assert.That(Eval("=INDEX(A:A,3)", sheet, workbook) as double?).IsEqualTo(30.0);

        // Data starting at row 5: the two conventions now DIVERGE. Absolute row 3 is A3 (genuinely blank —
        // this direct Expression.Evaluate path keeps blank AS blank, unlike the cell-boundary never-blank
        // rule GetCellValue applies); "the 3rd populated row" would have been A7 = 30.
        // ROW(INDEX(A:A,3)) = 3 pins the ADDRESS unambiguously, independent of what value sits there.
        var (workbookB, sheetB) = Sheet(("A5", Number(10)), ("A6", Number(20)), ("A7", Number(30)));

        await Assert.That(Eval("=INDEX(A:A,3)", sheetB, workbookB)).IsNull();
        await Assert.That(Eval("=ROW(INDEX(A:A,3))", sheetB, workbookB) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Offset_WholeColumnBase()
    {
        // OFFSET's open base is its ABSOLUTE first cell — row 1 of column A, sweep item 37 follow-up
        // ruling (a) — which happens to be A1 here regardless of population; offset (row 1, col 0) from
        // there is A2. (Coincides with "the populated box's own top-left" only because data starts at row
        // 1 — see Index_WholeColumn_ByAbsolutePosition for a fixture where the two conventions diverge.)
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
