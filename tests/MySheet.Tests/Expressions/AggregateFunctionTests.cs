using MySheet.Expressions;
using static MySheet.Expressions.Expression;

namespace MySheet.Tests.Expressions;

public class AggregateFunctionTests
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

    // --- Sum: the load-bearing bug fixes ---

    [Test]
    public async Task Sum_WithBinaryOperationArgument_DoesNotThrow()
    {
        var (workbook, _) = Sheet();

        var expr = Sum(Add(Number(1), Number(1)), Number(3));

        await Assert.That(expr.Compute(workbook) as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task Sum_WithChainedCellReference_Computes()
    {
        // A3 = SUM(A1, A2) = 3 ; SUM(A3, 1) = 4
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)));
        sheet["A3"] = Sum(Cell("A1", sheet), Cell("A2", sheet));

        var expr = Sum(Cell("A3", sheet), Number(1));

        await Assert.That(expr.Compute(workbook) as double?).IsEqualTo(4.0);
    }

    [Test]
    public async Task Sum_OverRange_Computes()
    {
        var (workbook, sheet) = Sheet(("A1", Number(1)), ("A2", Number(2)), ("A3", Number(3)));

        var expr = Sum(Range("A1", "A3", sheet));

        await Assert.That(expr.Compute(workbook) as double?).IsEqualTo(6.0);
    }

    [Test]
    public async Task Sum_DirectNonNumericText_IsValueError()
    {
        var (workbook, _) = Sheet();

        var expr = Sum(new MySheet.Expressions.StringValue("abc"), Number(1));

        await Assert.That(expr.Compute(workbook)).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Sum_ReferencedText_IsIgnored()
    {
        var (workbook, sheet) = Sheet(("A1", new MySheet.Expressions.StringValue("abc")));

        var expr = Sum(Cell("A1", sheet), Number(1));

        await Assert.That(expr.Compute(workbook) as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task Sum_PropagatesReferencedError()
    {
        var (workbook, sheet) = Sheet(("A1", Divide(Number(1), Number(0))));

        var expr = Sum(Cell("A1", sheet), Number(1));

        await Assert.That(expr.Compute(workbook)).IsEqualTo(ErrorValue.DivByZero);
    }

    // --- Average / Min / Max / Count ---

    [Test]
    public async Task Average_Computes()
    {
        var (workbook, _) = Sheet();

        await Assert.That(Average(Number(1), Number(3)).Compute(workbook) as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Average_OfEmpty_IsDivByZero()
    {
        var (workbook, _) = Sheet();

        await Assert.That(Average().Compute(workbook)).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Min_Computes()
    {
        var (workbook, _) = Sheet();

        await Assert.That(Min(Number(3), Number(1), Number(2)).Compute(workbook) as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task Max_Computes()
    {
        var (workbook, _) = Sheet();

        await Assert.That(Max(Number(3), Number(1), Number(2)).Compute(workbook) as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Min_OfEmpty_IsZero()
    {
        var (workbook, _) = Sheet();

        await Assert.That(Min().Compute(workbook) as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task Count_CountsOnlyNumbers()
    {
        var (workbook, sheet) = Sheet(
            ("A1", Number(1)),
            ("A2", new MySheet.Expressions.StringValue("x")),
            ("A3", Number(3)));

        var expr = Count(Range("A1", "A3", sheet));

        await Assert.That(expr.Compute(workbook) as double?).IsEqualTo(2.0);
    }
}
