using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class LookupFunctionTests
{
    private static object? Calc(string formula, params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = new NumberValue(value);
        }

        return ExpressionParser.Parse(formula, sheet).Compute(workbook);
    }

    [Test]
    public async Task Rows_CountsRowsInRange()
    {
        await Assert.That(Calc("=ROWS(A1:A3)") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=ROWS(A1:C2)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Row_ReturnsRowOfReference()
    {
        await Assert.That(Calc("=ROW(A5)") as double?).IsEqualTo(5.0);
        await Assert.That(Calc("=ROW(B2:B4)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Match_Exact()
    {
        await Assert
            .That(Calc("=MATCH(20,A1:A3,0)", ("A1", 10), ("A2", 20), ("A3", 30)) as double?)
            .IsEqualTo(2.0);
        await Assert
            .That(Calc("=MATCH(99,A1:A3,0)", ("A1", 10), ("A2", 20), ("A3", 30)))
            .IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Match_ApproximateAscending()
    {
        // Largest value <= 25 is 20, at position 2.
        await Assert
            .That(Calc("=MATCH(25,A1:A3,1)", ("A1", 10), ("A2", 20), ("A3", 30)) as double?)
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task Index_SingleColumn()
    {
        await Assert
            .That(Calc("=INDEX(A1:A3,2)", ("A1", 10), ("A2", 20), ("A3", 30)) as double?)
            .IsEqualTo(20.0);
    }

    [Test]
    public async Task Index_SingleRow_TreatsArgAsColumn()
    {
        await Assert
            .That(Calc("=INDEX(A1:C1,2)", ("A1", 7), ("B1", 8), ("C1", 9)) as double?)
            .IsEqualTo(8.0);
    }

    [Test]
    public async Task Index_TwoDimensional()
    {
        await Assert
            .That(Calc("=INDEX(A1:B2,2,2)", ("A1", 1), ("B1", 2), ("A2", 3), ("B2", 4)) as double?)
            .IsEqualTo(4.0);
    }
}
