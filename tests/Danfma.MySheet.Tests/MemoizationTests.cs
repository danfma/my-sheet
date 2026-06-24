using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

public class MemoizationTests
{
    [Test]
    public async Task ReferencedCell_IsComputedOnce()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        var calls = 0;
        workbook.RegisterFunction(
            "TICK",
            (_, _) =>
            {
                calls++;
                return 5.0;
            }
        );

        sheet["A1"] = ExpressionParser.Parse("=TICK()", sheet);
        sheet["B1"] = ExpressionParser.Parse("=A1+A1", sheet);

        var result = sheet["B1"].Compute(workbook) as double?;

        await Assert.That(result).IsEqualTo(10.0);
        // A1 is referenced twice but cached, so TICK runs once.
        await Assert.That(calls).IsEqualTo(1);
    }

    [Test]
    public async Task InvalidateCache_RefreshesAfterMutation()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new NumberValue(10);
        sheet["B1"] = ExpressionParser.Parse("=A1", sheet);

        await Assert.That(sheet["B1"].Compute(workbook) as double?).IsEqualTo(10.0);

        sheet["A1"] = new NumberValue(20);

        // The cache is explicit: without invalidation the old value is still served.
        await Assert.That(sheet["B1"].Compute(workbook) as double?).IsEqualTo(10.0);

        workbook.InvalidateCache();

        await Assert.That(sheet["B1"].Compute(workbook) as double?).IsEqualTo(20.0);
    }

    [Test]
    public async Task RangeAndDirectReference_ShareCache()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        var calls = 0;
        workbook.RegisterFunction(
            "TICK",
            (_, _) =>
            {
                calls++;
                return 5.0;
            }
        );

        sheet["A1"] = ExpressionParser.Parse("=TICK()", sheet);
        sheet["B1"] = ExpressionParser.Parse("=SUM(A1:A1)+A1", sheet);

        var result = sheet["B1"].Compute(workbook) as double?;

        await Assert.That(result).IsEqualTo(10.0);
        // A1 reached through the range expansion and the direct reference shares one cache entry.
        await Assert.That(calls).IsEqualTo(1);
    }

    [Test]
    public async Task CircularReference_ReturnsRefError_NotStackOverflow()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = ExpressionParser.Parse("=B1", sheet);
        sheet["B1"] = ExpressionParser.Parse("=A1", sheet);

        await Assert
            .That(ExpressionParser.Parse("=A1", sheet).Compute(workbook))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task RunWithLargeStack_HandlesDeepDependencyChain()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // A cumulative chain deep enough to overflow the default (~1 MB) stack.
        const int depth = 20000;
        sheet["A1"] = new NumberValue(1);
        for (var i = 2; i <= depth; i++)
        {
            sheet[$"A{i}"] = ExpressionParser.Parse($"=A{i - 1}+1", sheet);
        }

        var result =
            Workbook.RunWithLargeStack(() => sheet[$"A{depth}"].Compute(workbook)) as double?;

        await Assert.That(result).IsEqualTo((double)depth);
    }
}
