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
}
