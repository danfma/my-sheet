using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ConditionalTests
{
    private static object? Calc(string formula, params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return ExpressionParser.Parse(formula, sheet).Compute(workbook);
    }

    private static Expression ParseOnly(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet);
    }

    // --- IF ---

    [Test]
    public async Task If_TrueBranch()
    {
        await Assert.That(Calc("=IF(1>0,10,20)") as double?).IsEqualTo(10.0);
    }

    [Test]
    public async Task If_FalseBranch()
    {
        await Assert.That(Calc("=IF(0>1,10,20)") as double?).IsEqualTo(20.0);
    }

    [Test]
    public async Task If_TwoArgs_TrueReturnsValue()
    {
        await Assert.That(Calc("=IF(1>0,99)") as double?).IsEqualTo(99.0);
    }

    [Test]
    public async Task If_TwoArgs_FalseReturnsFalse()
    {
        await Assert.That(Calc("=IF(0>1,99)") as bool?).IsFalse();
    }

    [Test]
    public async Task If_ShortCircuits_UntakenBranch()
    {
        // A1 = 0; condition true -> returns 0; the false branch 1/A1 must NOT be evaluated.
        await Assert
            .That(Calc("=IF(A1=0,0,1/A1)", ("A1", new NumberValue(0))) as double?)
            .IsEqualTo(0.0);
    }

    [Test]
    public async Task If_TextCondition()
    {
        await Assert
            .That(
                Calc(
                    "=IF(A1=\"sim\",1,0)",
                    ("A1", new Danfma.MySheet.Expressions.StringValue("sim"))
                ) as double?
            )
            .IsEqualTo(1.0);
    }

    // --- Condition truthiness ---

    [Test]
    public async Task If_NonZeroNumberIsTrue()
    {
        await Assert.That(Calc("=IF(2,1,0)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task If_ZeroIsFalse()
    {
        await Assert.That(Calc("=IF(0,1,0)") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task If_TextConditionIsValueError()
    {
        await Assert.That(Calc("=IF(\"x\",1,0)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task If_ErrorConditionPropagates()
    {
        await Assert.That(Calc("=IF(1/0,1,0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    // --- AND / OR / NOT ---

    [Test]
    public async Task And_AllTrue()
    {
        await Assert.That(Calc("=AND(1>0,2>1)") as bool?).IsTrue();
    }

    [Test]
    public async Task And_OneFalse()
    {
        await Assert.That(Calc("=AND(1>0,2<1)") as bool?).IsFalse();
    }

    [Test]
    public async Task Or_OneTrue()
    {
        await Assert.That(Calc("=OR(1<0,2>1)") as bool?).IsTrue();
    }

    [Test]
    public async Task Not_InvertsCondition()
    {
        await Assert.That(Calc("=NOT(1>0)") as bool?).IsFalse();
    }

    // --- IFERROR ---

    [Test]
    public async Task IfError_ReturnsFallbackOnError()
    {
        await Assert.That(Calc("=IFERROR(1/0,-1)") as double?).IsEqualTo(-1.0);
    }

    [Test]
    public async Task IfError_ReturnsValueWhenNoError()
    {
        await Assert.That(Calc("=IFERROR(5,-1)") as double?).IsEqualTo(5.0);
    }

    [Test]
    public async Task IfError_ShortCircuits_FallbackNotEvaluatedWhenNoError()
    {
        // value is fine -> the fallback 1/0 must NOT be evaluated, so no error surfaces.
        await Assert.That(Calc("=IFERROR(5,1/0)") as double?).IsEqualTo(5.0);
    }

    // --- Arity errors throw ParseException ---

    [Test]
    public async Task If_TooFewArguments_Throws()
    {
        await Assert.That(() => ParseOnly("=IF(1)")).Throws<ParseException>();
    }

    [Test]
    public async Task If_TooManyArguments_Throws()
    {
        await Assert.That(() => ParseOnly("=IF(1,2,3,4)")).Throws<ParseException>();
    }

    [Test]
    public async Task Not_WrongArity_Throws()
    {
        await Assert.That(() => ParseOnly("=NOT(1,2)")).Throws<ParseException>();
    }

    [Test]
    public async Task IfError_WrongArity_Throws()
    {
        await Assert.That(() => ParseOnly("=IFERROR(1)")).Throws<ParseException>();
    }

    // --- Regression: variadic functions and unknown names still behave ---

    [Test]
    public async Task Sum_StillAcceptsAnyArity()
    {
        await Assert.That(Calc("=SUM()") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task UnknownFunction_StillNameError()
    {
        await Assert.That(Calc("=FOO(1)")).IsEqualTo(ErrorValue.Name);
    }
}
