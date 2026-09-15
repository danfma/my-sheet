using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class NegativeZeroTests
{
    [Test]
    [Arguments("=-SUM(A1:A2)")]
    [Arguments("=-0")]
    [Arguments("=0*-1")]
    [Arguments("=-A1")]
    [Arguments("=A1*-1")]
    [Arguments("=ROUND(-0.0001,2)")]
    [Arguments("=VALUE(\"-0\")")]
    [Arguments("=ROUNDDOWN(-0.01,1)")]
    [Arguments("=CEILING(-0.4,1)")]
    [Arguments("=INT(-0.0)")]
    [Arguments("=MOD(-SUM(A1:A2),3)")]
    [Arguments("=MAX(-A1:A2)")]
    [Arguments("=INDEX(-A1:A2,1)")]
    public async Task ComputedNumericZero_HasNoNegativeSign(string formula)
    {
        var value = Evaluate(formula);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0d);
        await Assert.That(double.IsNegative(value.ToDouble())).IsFalse();
    }

    [Test]
    [Arguments("=TEXT(-SUM(A1:A2),\"0.0\")", "0.0")]
    [Arguments("=TEXT(-SUM(A1:A2),\"0.0;(0.0)\")", "0.0")]
    [Arguments("=-SUM(A1:A2)&\"\"", "0")]
    [Arguments("=FIXED(-SUM(A1:A2),1)", "0.0")]
    [Arguments("=DOLLAR(-SUM(A1:A2))", "$0.00")]
    [Arguments("=TEXT(INDEX(-A1:A2,2),\"0.0\")", "0.0")]
    [Arguments("=LET(x,-A1,TEXT(x,\"0.0\"))", "0.0")]
    public async Task Consumers_SeeNormalizedComputedZero(string formula, string expected)
    {
        await Assert.That(Evaluate(formula).ToText()).IsEqualTo(expected);
    }

    [Test]
    public async Task DefinedName_SeesNormalizedComputedZero()
    {
        var workbook = CreateWorkbook(out var sheet);
        workbook.DefineName("NegZero", new NumberValue(-0d));

        var result = ExpressionParser.Parse("=TEXT(NegZero,\"0.0\")", sheet).Evaluate(workbook);

        await Assert.That(result.ToText()).IsEqualTo("0.0");
    }

    [Test]
    public async Task PublicRead_NormalizesLiteralNegativeZero()
    {
        var workbook = CreateWorkbook(out var sheet);
        sheet["A3"] = new NumberValue(-0d);
        sheet["A4"] = ExpressionParser.Parse("=A3", sheet);
        sheet["A5"] = ExpressionParser.Parse("=TEXT(A3,\"0.0\")", sheet);

        var value = workbook.GetCellValue("Sheet1", "A4").ToDouble();
        await Assert.That(value).IsEqualTo(0d);
        await Assert.That(double.IsNegative(value)).IsFalse();
        await Assert.That(workbook.GetCellValue("Sheet1", "A5").ToText()).IsEqualTo("0.0");
    }

    [Test]
    public async Task NonZeroNegativeValues_ThatDisplayAsZero_KeepTheirSign()
    {
        await Assert.That(Evaluate("=TEXT(-0.0001,\"0.0\")").ToText()).IsEqualTo("0.0");
        await Assert.That(Evaluate("=TEXT(-0.04,\"0.0\")").ToText()).IsEqualTo("0.0");
    }

    [Test]
    public async Task Text_UsesOriginalSignForSections_AndRoundedMagnitudeForSingleSectionMinus()
    {
        var values = new[] { "-0.4", "-0.04", "-0.0001", "-0.5", "-0.6", "-SUM(A1:A2)", "0.4" };
        var formats = new[]
        {
            "0",
            "0.0",
            "00.00",
            "#,##0",
            "0%",
            "0.0,",
            "$0.00",
            "0 \"units\"",
            "\"-\"0",
            "0.0E+0",
            "0;(0)",
            "0.0;-0.0",
            "0;\"neg\"",
            "0;(0);\"zero\"",
            "0.0;(0.0);\"z\"",
        };
        var expected = new[]
        {
            new[]
            {
                "0",
                "-0.4",
                "-00.40",
                "0",
                "-40%",
                "0.0",
                "-$0.40",
                "0 units",
                "-0",
                "-4.0E-1",
                "(0)",
                "-0.4",
                "neg",
                "(0)",
                "(0.4)",
            },
            new[]
            {
                "0",
                "0.0",
                "-00.04",
                "0",
                "-4%",
                "0.0",
                "-$0.04",
                "0 units",
                "-0",
                "-4.0E-2",
                "(0)",
                "-0.0",
                "neg",
                "(0)",
                "(0.0)",
            },
            new[]
            {
                "0",
                "0.0",
                "00.00",
                "0",
                "0%",
                "0.0",
                "$0.00",
                "0 units",
                "-0",
                "-1.0E-4",
                "(0)",
                "-0.0",
                "neg",
                "(0)",
                "(0.0)",
            },
            new[]
            {
                "-1",
                "-0.5",
                "-00.50",
                "-1",
                "-50%",
                "0.0",
                "-$0.50",
                "-1 units",
                "--1",
                "-5.0E-1",
                "(1)",
                "-0.5",
                "neg",
                "(1)",
                "(0.5)",
            },
            new[]
            {
                "-1",
                "-0.6",
                "-00.60",
                "-1",
                "-60%",
                "0.0",
                "-$0.60",
                "-1 units",
                "--1",
                "-6.0E-1",
                "(1)",
                "-0.6",
                "neg",
                "(1)",
                "(0.6)",
            },
            new[]
            {
                "0",
                "0.0",
                "00.00",
                "0",
                "0%",
                "0.0",
                "$0.00",
                "0 units",
                "-0",
                "0.0E+0",
                "0",
                "0.0",
                "0",
                "zero",
                "z",
            },
            new[]
            {
                "0",
                "0.4",
                "00.40",
                "0",
                "40%",
                "0.0",
                "$0.40",
                "0 units",
                "-0",
                "4.0E-1",
                "0",
                "0.4",
                "0",
                "0",
                "0.4",
            },
        };

        for (var valueIndex = 0; valueIndex < values.Length; valueIndex++)
        {
            for (var formatIndex = 0; formatIndex < formats.Length; formatIndex++)
            {
                var formula =
                    $"=TEXT({values[valueIndex]},\"{formats[formatIndex].Replace("\"", "\"\"")}\")";
                await Assert
                    .That(Evaluate(formula).ToText())
                    .IsEqualTo(expected[valueIndex][formatIndex]);
            }
        }
    }

    [Test]
    [Arguments("=SQRT(-SUM(A1:A2))")]
    [Arguments("=POWER(-SUM(A1:A2),3)")]
    [Arguments("=(-SUM(A1:A2))^3")]
    [Arguments("=PRODUCT(-1,A1)")]
    [Arguments("=TRUNC(-0.5)")]
    [Arguments("=CEILING.MATH(-0.5)")]
    [Arguments("=MIN(-A1:A2)")]
    [Arguments("=SUM(-A1:A2)")]
    [Arguments("=ABS(-SUM(A1:A2))")]
    [Arguments("=N(-SUM(A1:A2))")]
    [Arguments("=IF(TRUE,-A1)")]
    [Arguments("=CHOOSE(1,-A1)")]
    public async Task AdversarialNumericZero_HasNoNegativeSign(string formula)
    {
        var value = Evaluate(formula);

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0d);
        await Assert.That(double.IsNegative(value.ToDouble())).IsFalse();
    }

    [Test]
    public async Task CountIf_TreatsNegativeZeroAsZero()
    {
        await Assert.That(Evaluate("=COUNTIF(A1:A2,-0)").ToDouble()).IsEqualTo(2d);
    }

    [Test]
    public async Task ZeroDivisionAndEqualityGuards_StayUnchanged()
    {
        await Assert
            .That(Evaluate("=1/-SUM(A1:A2)").TryGetError(out var error) ? error : default)
            .IsEqualTo(Error.DivZero);
        await Assert.That(Evaluate("=-SUM(A1:A2)=0").ToBoolean()).IsTrue();
    }

    private static ComputedValue Evaluate(string formula)
    {
        var workbook = CreateWorkbook(out var sheet);
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook);
    }

    private static Workbook CreateWorkbook(out Sheet sheet)
    {
        var workbook = new Workbook();
        sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(0d);
        sheet["A2"] = new NumberValue(0d);
        return workbook;
    }
}
