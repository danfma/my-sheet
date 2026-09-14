using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ConditionalAggregationTests
{
    private static object? Calc(string formula, params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static (string, Expression) N(string id, double v) => (id, new NumberValue(v));

    private static (string, Expression) T(string id, string v) =>
        (id, new Danfma.MySheet.Expressions.StringValue(v));

    [Test]
    public async Task CountA_CountsNonBlank()
    {
        await Assert
            .That(Calc("=COUNTA(A1:A3)", N("A1", 1), T("A2", "x")) as double?)
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task CountBlank_CountsEmpty()
    {
        await Assert
            .That(Calc("=COUNTBLANK(A1:A3)", N("A1", 1), T("A2", "x")) as double?)
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task CountIf_NumericCriteria()
    {
        await Assert
            .That(Calc("=COUNTIF(A1:A3,\">1\")", N("A1", 1), N("A2", 2), N("A3", 3)) as double?)
            .IsEqualTo(2.0);
        await Assert
            .That(Calc("=COUNTIF(A1:A3,2)", N("A1", 1), N("A2", 2), N("A3", 3)) as double?)
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task CountIf_TextWildcard()
    {
        await Assert
            .That(Calc("=COUNTIF(A1:A2,\"a*\")", T("A1", "apple"), T("A2", "banana")) as double?)
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task SumIf_WithAndWithoutSumRange()
    {
        await Assert
            .That(Calc("=SUMIF(A1:A3,\">1\")", N("A1", 1), N("A2", 2), N("A3", 3)) as double?)
            .IsEqualTo(5.0);

        await Assert
            .That(
                Calc(
                    "=SUMIF(A1:A3,\">1\",B1:B3)",
                    N("A1", 1),
                    N("A2", 2),
                    N("A3", 3),
                    N("B1", 10),
                    N("B2", 20),
                    N("B3", 30)
                ) as double?
            )
            .IsEqualTo(50.0);
    }

    [Test]
    public async Task SumIf_ResizesSumRangeFromItsTopLeft()
    {
        // Aspose.Cells 26.7.0 gives 4 for B1, B1:B2 and B1:C1: SUMIF resizes from the supplied
        // sum_range's top-left to the 3x1 criteria-range shape. MySheet previously truncated to the
        // supplied shape (1 for B1, 3 for B1:B2 and 1 for B1:C1 on this discriminating fixture).
        foreach (var sumRange in new[] { "B1", "B1:B2", "B1:C1" })
        {
            await Assert
                .That(
                    Calc(
                        $"=SUMIF(A1:A3,\">0\",{sumRange})",
                        N("A1", 5),
                        N("A2", 0),
                        N("A3", 9),
                        N("B1", 1),
                        N("B2", 2),
                        N("B3", 3),
                        N("C1", 10)
                    ) as double?
                )
                .IsEqualTo(4.0);
        }

        // B3 is resized to B3:B5. The cells beyond the used area are ordinary blanks, so only B3 adds 3.
        await Assert
            .That(
                Calc("=SUMIF(A1:A3,\">0\",B3)", N("A1", 5), N("A2", 0), N("A3", 9), N("B3", 3))
                    as double?
            )
            .IsEqualTo(3.0);
    }

    [Test]
    public async Task AverageIf_ResizesAverageRangeFromItsTopLeft()
    {
        // Aspose.Cells 26.7.0 gives 2; MySheet previously read only B1 and gave 1.
        await Assert
            .That(
                Calc(
                    "=AVERAGEIF(A1:A3,\">0\",B1:B2)",
                    N("A1", 5),
                    N("A2", 0),
                    N("A3", 9),
                    N("B1", 1),
                    N("B2", 2),
                    N("B3", 3)
                ) as double?
            )
            .IsEqualTo(2.0);
    }

    [Test]
    public async Task CountIfs_MultipleCriteria()
    {
        await Assert
            .That(
                Calc("=COUNTIFS(A1:A3,\">1\",A1:A3,\"<3\")", N("A1", 1), N("A2", 2), N("A3", 3))
                    as double?
            )
            .IsEqualTo(1.0);
    }

    [Test]
    public async Task SumIfs_MultipleCriteria()
    {
        await Assert
            .That(
                Calc(
                    "=SUMIFS(B1:B3,A1:A3,\">1\")",
                    N("A1", 1),
                    N("A2", 2),
                    N("A3", 3),
                    N("B1", 10),
                    N("B2", 20),
                    N("B3", 30)
                ) as double?
            )
            .IsEqualTo(50.0);
    }
}
