using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// The Layer-1 <see cref="SheetStructuralIndex"/> (whole-column scale): it is built once per cache epoch
/// and reused, dropped by <see cref="Workbook.InvalidateCache"/>, survives <see cref="Workbook.Recalculate"/>,
/// enumerates in column-then-row order, and builds its column and row maps independently.
/// </summary>
public class StructuralIndexTests
{
    private static (Workbook Workbook, Sheet Sheet) Sheet(params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return (workbook, sheet);
    }

    private static void Eval(string formula, Sheet sheet, Workbook workbook) =>
        _ = ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

    private static SheetStructuralIndex Index(Workbook workbook) =>
        workbook.GetStructuralIndex("Sheet1", workbook["Sheet1"]);

    [Test]
    public async Task Index_BuiltOnce_ReusedAcrossReads()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        Eval("=SUM(A:A)", sheet, workbook);
        Eval("=COUNTA(A:A)", sheet, workbook);
        Eval("=MAX(A:A)", sheet, workbook);

        await Assert.That(Index(workbook).ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task Index_DroppedByInvalidateCache()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Eval("=SUM(A:A)", sheet, workbook);
        var before = Index(workbook);

        workbook.InvalidateCache();
        var after = Index(workbook);

        await Assert.That(ReferenceEquals(before, after)).IsFalse();
    }

    [Test]
    public async Task Index_SurvivesRecalculate()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        Eval("=SUM(A:A)", sheet, workbook);
        var before = Index(workbook);

        workbook.Recalculate();
        var after = Index(workbook);

        await Assert.That(ReferenceEquals(before, after)).IsTrue();
        // The surviving instance is not rebuilt on the next read.
        Eval("=SUM(A:A)", sheet, workbook);
        await Assert.That(after.ColumnBuildCount).IsEqualTo(1);
    }

    [Test]
    public async Task PopulatedIds_ColumnThenRowOrder_RegardlessOfInsertionOrder()
    {
        // Insert deliberately out of spatial order; the index must still yield column-then-row.
        var (workbook, _) = Sheet(
            ("B1", Number(10)),
            ("A3", Number(3)),
            ("A1", Number(1)),
            ("A2", Number(2)),
            ("B2", Number(20))
        );

        var open = OpenRangeReference.Create(1, 2, null, null, "Sheet1"); // A:B
        var ids = open.PopulatedIds(new EvaluationContext(workbook)).ToList();

        await Assert.That(string.Join(",", ids)).IsEqualTo("A1,A2,A3,B1,B2");
    }

    [Test]
    public async Task WholeRow_BuildsRowMapOnly_NotColumnMap()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("B1", Number(2)), ("A2", Number(999)));

        Eval("=SUM(1:1)", sheet, workbook);

        var index = Index(workbook);
        await Assert.That(index.RowBuildCount).IsEqualTo(1);
        await Assert.That(index.ColumnBuildCount).IsEqualTo(0);
    }

    [Test]
    public async Task WholeColumn_BuildsColumnMapOnly_NotRowMap()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("B1", Number(2)), ("A2", Number(3)));

        Eval("=SUM(A:A)", sheet, workbook);

        var index = Index(workbook);
        await Assert.That(index.ColumnBuildCount).IsEqualTo(1);
        await Assert.That(index.RowBuildCount).IsEqualTo(0);
    }
}
