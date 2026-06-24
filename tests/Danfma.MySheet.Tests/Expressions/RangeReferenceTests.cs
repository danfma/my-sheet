using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

public class RangeReferenceTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["B1"] = Number(3);
        sheet["B2"] = Number(4);

        return (workbook, sheet);
    }

    [Test]
    public async Task Compute_InScalarContext_ReturnsValueError()
    {
        var (workbook, sheet) = Grid();

        await Assert
            .That(Range("A1", "B2", sheet).Compute(workbook))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Expand_YieldsEveryCellInRectangle()
    {
        var (workbook, sheet) = Grid();

        var cells = Range("A1", "B2", sheet).Expand(workbook).ToList();

        await Assert.That(cells.Count).IsEqualTo(4);
        await Assert.That(cells.Sum(e => (double)e.Compute(workbook)!)).IsEqualTo(10.0);
    }

    [Test]
    public async Task Expand_NormalizesReversedRange()
    {
        var (workbook, sheet) = Grid();

        var cells = Range("B2", "A1", sheet).Expand(workbook).ToList();

        await Assert.That(cells.Count).IsEqualTo(4);
    }

    [Test]
    public async Task Expand_HandlesSingleCellRange()
    {
        var (workbook, sheet) = Grid();

        var cells = Range("A1", "A1", sheet).Expand(workbook).ToList();

        await Assert.That(cells.Count).IsEqualTo(1);
    }
}
