using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Whole-column (<c>A:A</c>), whole-row (<c>1:1</c>) and one-sided open (<c>A2:A</c>, <c>A:A10</c>,
/// <c>A1:C</c>) references. Semantics are "the populated cells within the limits"; blanks contribute 0,
/// so aggregation matches Excel.
/// </summary>
public class WholeColumnReferenceTests
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

    private static double? Eval(string formula, Sheet sheet, Workbook workbook) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject() as double?;

    [Test]
    public async Task Parse_WholeColumn_ProducesOpenRangeReference()
    {
        var sheet = new Sheet { Name = "Sheet1" };

        var expression = ExpressionParser.Parse("=A:A", sheet);

        await Assert.That(expression is OpenRangeReference).IsTrue();
    }

    [Test]
    public async Task Parse_FullyBoundedRange_StaysRangeReference()
    {
        var sheet = new Sheet { Name = "Sheet1" };

        var expression = ExpressionParser.Parse("=A1:B2", sheet);

        await Assert.That(expression is RangeReference).IsTrue();
    }

    [Test]
    public async Task Sum_WholeColumn_SumsPopulatedCells()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        await Assert.That(Eval("=SUM(A:A)", sheet, workbook)).IsEqualTo(6.0);
    }

    [Test]
    public async Task Sum_WholeColumn_IgnoresOtherColumns()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("B1", Number(100)));

        await Assert.That(Eval("=SUM(A:A)", sheet, workbook)).IsEqualTo(3.0);
    }

    [Test]
    public async Task Average_WholeColumn_AveragesPopulatedCells()
    {
        var (workbook, sheet) = Sheet(("A1", Number(2)), ("A2", Number(4)), ("A3", Number(6)));

        await Assert.That(Eval("=AVERAGE(A:A)", sheet, workbook)).IsEqualTo(4.0);
    }

    [Test]
    public async Task CountA_WholeColumn_CountsPopulatedCells()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A5", String("x")), ("A100", Number(9)));

        await Assert.That(Eval("=COUNTA(A:A)", sheet, workbook)).IsEqualTo(3.0);
    }

    [Test]
    public async Task Max_MultiColumn_UsesAllColumnsInRange()
    {
        var (workbook, sheet) = Sheet(
            ("A1", Number(1)),
            ("B1", Number(5)),
            ("C1", Number(3)),
            ("D1", Number(99))
        );

        await Assert.That(Eval("=MAX(A:C)", sheet, workbook)).IsEqualTo(5.0);
    }

    [Test]
    public async Task Sum_WholeRow_SumsPopulatedCells()
    {
        var (workbook, sheet) = Sheet(("A1", Number(10)), ("B1", Number(20)), ("A2", Number(999)));

        await Assert.That(Eval("=SUM(1:1)", sheet, workbook)).IsEqualTo(30.0);
    }

    [Test]
    public async Task Sum_MultiRow_SumsAllRowsInRange()
    {
        var (workbook, sheet) = Sheet(
            ("A1", Number(1)),
            ("A2", Number(2)),
            ("A5", Number(5)),
            ("A6", Number(100))
        );

        await Assert.That(Eval("=SUM(1:5)", sheet, workbook)).IsEqualTo(8.0);
    }

    [Test]
    public async Task SumIf_WholeColumn_Filters()
    {
        var (workbook, sheet) = Sheet(
            ("A1", Number(3)),
            ("A2", Number(7)),
            ("A3", Number(10))
        );

        await Assert.That(Eval("=SUMIF(A:A,\">5\")", sheet, workbook)).IsEqualTo(17.0);
    }

    [Test]
    public async Task Sum_CrossSheetWholeColumn()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        var main = workbook.Sheets.Add("Main");
        data["A1"] = Number(4);
        data["A2"] = Number(6);

        await Assert.That(Eval("=SUM(Data!A:A)", main, workbook)).IsEqualTo(10.0);
    }

    // --- Mixed one-sided references ---

    [Test]
    public async Task Sum_MixedLowerRowBound_IgnoresAboveStart()
    {
        // A2:A covers column A from row 2 downward; A1 is excluded.
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        await Assert.That(Eval("=SUM(A2:A)", sheet, workbook)).IsEqualTo(5.0);
    }

    [Test]
    public async Task Sum_MixedUpperRowBound_IgnoresBelowEnd()
    {
        // A:A10 covers column A up to row 10; A11 is excluded.
        var (workbook, sheet) = Sheet(("A5", Number(5)), ("A10", Number(10)), ("A11", Number(1000)));

        await Assert.That(Eval("=SUM(A:A10)", sheet, workbook)).IsEqualTo(15.0);
    }

    [Test]
    public async Task Sum_MixedColumnsFromRow_IgnoresAboveStart()
    {
        // A1:C covers columns A..C from row 1 downward.
        var (workbook, sheet) = Sheet(
            ("A1", Number(1)),
            ("B2", Number(2)),
            ("C3", Number(3)),
            ("D1", Number(999))
        );

        await Assert.That(Eval("=SUM(A1:C)", sheet, workbook)).IsEqualTo(6.0);
    }

    [Test]
    public async Task Sum_SparseColumn_TouchesOnlyPopulatedCells()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A100", Number(2)));

        await Assert.That(Eval("=SUM(A:A)", sheet, workbook)).IsEqualTo(3.0);
    }

    [Test]
    public async Task Sum_EmptyColumn_IsZero()
    {
        var (workbook, sheet) = Sheet(("B1", Number(5)));

        await Assert.That(Eval("=SUM(A:A)", sheet, workbook)).IsEqualTo(0.0);
    }

    [Test]
    public async Task Sum_AbsoluteWholeColumn_NormalizesToRelative()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));

        await Assert.That(Eval("=SUM($A:$A)", sheet, workbook)).IsEqualTo(3.0);
    }
}
