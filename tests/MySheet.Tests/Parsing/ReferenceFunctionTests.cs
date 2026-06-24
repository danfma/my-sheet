using MySheet.Expressions;
using MySheet.Parsing;

namespace MySheet.Tests.Parsing;

public class ReferenceFunctionTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid(params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return (workbook, sheet);
    }

    private static Expression N(double v) => new NumberValue(v);
    private static Expression T(string v) => new MySheet.Expressions.StringValue(v);

    [Test]
    public async Task Row_NoArgument_UsesCurrentCell()
    {
        var (workbook, sheet) = Grid();
        sheet["A5"] = ExpressionParser.Parse("=ROW()", sheet);

        // Reaching A5 through a reference sets the current cell to A5.
        await Assert.That(ExpressionParser.Parse("=A5", sheet).Compute(workbook) as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task VLookup_Exact()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)), ("B1", T("a")),
            ("A2", N(2)), ("B2", T("b")),
            ("A3", N(3)), ("B3", T("c")));

        await Assert.That(ExpressionParser.Parse("=VLOOKUP(2,A1:B3,2,FALSE)", sheet).Compute(workbook) as string)
            .IsEqualTo("b");
        await Assert.That(ExpressionParser.Parse("=VLOOKUP(99,A1:B3,2,FALSE)", sheet).Compute(workbook))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task VLookup_Approximate()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)), ("B1", T("a")),
            ("A2", N(2)), ("B2", T("b")),
            ("A3", N(3)), ("B3", T("c")));

        await Assert.That(ExpressionParser.Parse("=VLOOKUP(2.5,A1:B3,2,TRUE)", sheet).Compute(workbook) as string)
            .IsEqualTo("b");
    }

    [Test]
    public async Task XLookup_ExactAndNotFound()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)), ("B1", T("a")),
            ("A2", N(2)), ("B2", T("b")),
            ("A3", N(3)), ("B3", T("c")));

        await Assert.That(ExpressionParser.Parse("=XLOOKUP(2,A1:A3,B1:B3)", sheet).Compute(workbook) as string)
            .IsEqualTo("b");
        await Assert.That(ExpressionParser.Parse("=XLOOKUP(99,A1:A3,B1:B3,\"none\")", sheet).Compute(workbook) as string)
            .IsEqualTo("none");
        await Assert.That(ExpressionParser.Parse("=XLOOKUP(99,A1:A3,B1:B3)", sheet).Compute(workbook))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Offset_ScalarCell()
    {
        var (workbook, sheet) = Grid(("A1", N(10)), ("A2", N(20)), ("A3", N(30)), ("B1", N(5)));

        await Assert.That(ExpressionParser.Parse("=OFFSET(A1,2,0)", sheet).Compute(workbook) as double?)
            .IsEqualTo(30.0);
        await Assert.That(ExpressionParser.Parse("=OFFSET(A1,0,1)", sheet).Compute(workbook) as double?)
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task Offset_MultiCell_FeedsAggregation()
    {
        var (workbook, sheet) = Grid(("A1", N(10)), ("A2", N(20)), ("A3", N(30)));

        // OFFSET(A1,0,0,3,1) is the range A1:A3, so SUM over it is 60.
        await Assert.That(ExpressionParser.Parse("=SUM(OFFSET(A1,0,0,3,1))", sheet).Compute(workbook) as double?)
            .IsEqualTo(60.0);
    }

    [Test]
    public async Task XLookup_ApproximateModes()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)), ("B1", T("a")),
            ("A2", N(2)), ("B2", T("b")),
            ("A3", N(3)), ("B3", T("c")));

        // Omitted if_not_found (,,) then match_mode -1 (next smaller) / 1 (next larger).
        await Assert.That(ExpressionParser.Parse("=XLOOKUP(2.5,A1:A3,B1:B3,,-1)", sheet).Compute(workbook) as string)
            .IsEqualTo("b");
        await Assert.That(ExpressionParser.Parse("=XLOOKUP(2.5,A1:A3,B1:B3,,1)", sheet).Compute(workbook) as string)
            .IsEqualTo("c");
    }

    [Test]
    public async Task XLookup_Wildcard()
    {
        var (workbook, sheet) = Grid(("A1", T("apple")), ("A2", T("banana")), ("B1", N(1)), ("B2", N(2)));

        await Assert.That(ExpressionParser.Parse("=XLOOKUP(\"a*\",A1:A2,B1:B2,,2)", sheet).Compute(workbook) as double?)
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task XLookup_BinarySearchModesReturnCorrectResults()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)), ("B1", T("a")),
            ("A2", N(2)), ("B2", T("b")),
            ("A3", N(3)), ("B3", T("c")));

        // search_mode 2 / -2 (binary) are not optimized but must still return the correct match.
        await Assert.That(ExpressionParser.Parse("=XLOOKUP(2,A1:A3,B1:B3,,0,2)", sheet).Compute(workbook) as string)
            .IsEqualTo("b");
        await Assert.That(ExpressionParser.Parse("=XLOOKUP(2,A1:A3,B1:B3,,0,-2)", sheet).Compute(workbook) as string)
            .IsEqualTo("b");
    }

    [Test]
    public async Task XLookup_ReverseSearchFindsLast()
    {
        var (workbook, sheet) = Grid(
            ("A1", N(1)), ("A2", N(2)), ("A3", N(2)), ("A4", N(3)),
            ("B1", T("a")), ("B2", T("b")), ("B3", T("c")), ("B4", T("d")));

        await Assert.That(ExpressionParser.Parse("=XLOOKUP(2,A1:A4,B1:B4,,0,-1)", sheet).Compute(workbook) as string)
            .IsEqualTo("c");
    }
}
