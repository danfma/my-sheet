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

        var result = sheet["B1"].Evaluate(workbook).AsObject() as double?;

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

        await Assert.That(sheet["B1"].Evaluate(workbook).AsObject() as double?).IsEqualTo(10.0);

        sheet["A1"] = new NumberValue(20);

        // The cache is explicit: without invalidation the old value is still served.
        await Assert.That(sheet["B1"].Evaluate(workbook).AsObject() as double?).IsEqualTo(10.0);

        workbook.InvalidateCache();

        await Assert.That(sheet["B1"].Evaluate(workbook).AsObject() as double?).IsEqualTo(20.0);
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

        var result = sheet["B1"].Evaluate(workbook).AsObject() as double?;

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
            .That(ExpressionParser.Parse("=A1", sheet).Evaluate(workbook).AsObject())
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
            Workbook.RunWithLargeStack(() => sheet[$"A{depth}"].Evaluate(workbook).AsObject())
            as double?;

        await Assert.That(result).IsEqualTo((double)depth);
    }

    [Test]
    public async Task ComputeAll_EvaluatesEveryCellEagerlyAndMemoizes()
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
        sheet["C1"] = ExpressionParser.Parse("=B1*2", sheet);

        // Nothing evaluated yet.
        await Assert.That(calls).IsEqualTo(0);

        workbook.ComputeAll();

        // Every cell computed exactly once (memoized), regardless of sweep order.
        await Assert.That(calls).IsEqualTo(1);
        await Assert.That(workbook.GetCellValue("Sheet1", "C1").AsObject()).IsEqualTo(20.0);

        // A second ComputeAll is all cache hits — no recomputation.
        workbook.ComputeAll();
        await Assert.That(calls).IsEqualTo(1);
    }

    [Test]
    public async Task ComputeAll_HandlesDeepChainViaLargeStack()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        const int depth = 20000;
        sheet["A1"] = new NumberValue(1);
        for (var i = 2; i <= depth; i++)
        {
            sheet[$"A{i}"] = ExpressionParser.Parse($"=A{i - 1}+1", sheet);
        }

        // A plain sweep on the calling thread would overflow; ComputeAll runs on a large stack.
        workbook.ComputeAll();

        await Assert
            .That(workbook.GetCellValue("Sheet1", $"A{depth}").AsObject())
            .IsEqualTo((double)depth);
    }
}
