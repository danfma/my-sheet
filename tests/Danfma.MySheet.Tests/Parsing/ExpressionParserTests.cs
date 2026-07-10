using System.Linq;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ExpressionParserTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid(params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = new NumberValue(value);
        }

        return (workbook, sheet);
    }

    private static object? Calc(string formula)
    {
        var (workbook, sheet) = Grid();

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    // --- Precedence / associativity ---

    [Test]
    public async Task MultiplicationBeforeAddition()
    {
        await Assert.That(Calc("=1+2*3") as double?).IsEqualTo(7.0);
    }

    [Test]
    public async Task Concatenation_Operator()
    {
        await Assert.That(Calc("=\"a\"&\"b\"") as string).IsEqualTo("ab");
        await Assert.That(Calc("=1&2") as string).IsEqualTo("12");
        // '+' binds tighter than '&'
        await Assert.That(Calc("=\"x\"&1+1") as string).IsEqualTo("x2");
        // '&' binds tighter than '='
        await Assert.That(Calc("=\"a\"&\"b\"=\"ab\"") as bool?).IsTrue();
    }

    [Test]
    public async Task Percent_PostfixOperator()
    {
        await Assert.That(Calc("=50%") as double?).IsEqualTo(0.5);
        await Assert.That(Calc("=200%") as double?).IsEqualTo(2.0);
        // '%' binds tighter than + and ^
        await Assert.That(Calc("=100+50%") as double?).IsEqualTo(100.5);
        await Assert.That(Calc("=2^200%") as double?).IsEqualTo(4.0);
    }

    [Test]
    public async Task Parentheses_OverridePrecedence()
    {
        await Assert.That(Calc("=(1+2)*3") as double?).IsEqualTo(9.0);
    }

    [Test]
    public async Task Subtraction_IsLeftAssociative()
    {
        await Assert.That(Calc("=2-3-4") as double?).IsEqualTo(-5.0);
    }

    [Test]
    public async Task Division_IsLeftAssociative()
    {
        await Assert.That(Calc("=10/2/5") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task PowerBeforeMultiplication()
    {
        await Assert.That(Calc("=2*3^2") as double?).IsEqualTo(18.0);
    }

    // --- Power / unary (Excel quirk) ---

    [Test]
    public async Task UnaryMinus_BindsTighterThanPower()
    {
        // Excel: -2^2 == 4 (parsed as (-2)^2), not -4
        await Assert.That(Calc("=-2^2") as double?).IsEqualTo(4.0);
    }

    [Test]
    public async Task Power_AllowsUnaryOnRightOperand()
    {
        await Assert.That(Calc("=2^-2") as double?).IsEqualTo(0.25);
    }

    [Test]
    public async Task Power_IsRightAssociative()
    {
        await Assert.That(Calc("=2^3^2") as double?).IsEqualTo(512.0);
    }

    [Test]
    public async Task StackedUnaryMinus_CancelsOut()
    {
        await Assert.That(Calc("=--2") as double?).IsEqualTo(2.0);
    }

    // --- Division by zero ---

    [Test]
    public async Task DivisionByZero_IsError()
    {
        await Assert.That(Calc("=1/0")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task DivisionByZero_PropagatesThroughArithmetic()
    {
        await Assert.That(Calc("=1/0+5")).IsEqualTo(ErrorValue.DivByZero);
    }

    // --- Comparators ---

    [Test]
    public async Task Comparators_ReturnBooleans()
    {
        await Assert.That(Calc("=1<2") as bool?).IsTrue();
        await Assert.That(Calc("=2<=2") as bool?).IsTrue();
        await Assert.That(Calc("=1<>2") as bool?).IsTrue();
        await Assert.That(Calc("=2=2") as bool?).IsTrue();
    }

    // --- Functions and references ---

    [Test]
    public async Task Sum_OfCells()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2));

        var result =
            ExpressionParser.Parse("=SUM(A1,A2)", sheet).Evaluate(workbook).AsObject() as double?;

        await Assert.That(result).IsEqualTo(3.0);
    }

    [Test]
    public async Task CellReferences_AreCaseInsensitive()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2));

        var result =
            ExpressionParser.Parse("=sum(a1,a2)", sheet).Evaluate(workbook).AsObject() as double?;

        await Assert.That(result).IsEqualTo(3.0);
    }

    [Test]
    public async Task Average_OfCells()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2));

        var result =
            ExpressionParser.Parse("=AVERAGE(A1,A2)", sheet).Evaluate(workbook).AsObject()
            as double?;

        await Assert.That(result).IsEqualTo(1.5);
    }

    [Test]
    public async Task Sum_OverRange()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));

        var result =
            ExpressionParser.Parse("=SUM(A1:A3)", sheet).Evaluate(workbook).AsObject() as double?;

        await Assert.That(result).IsEqualTo(6.0);
    }

    [Test]
    public async Task Sum_OverRange_MixedWithScalar()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3), ("B1", 10));

        var result =
            ExpressionParser.Parse("=SUM(A1:A3, B1)", sheet).Evaluate(workbook).AsObject()
            as double?;

        await Assert.That(result).IsEqualTo(16.0);
    }

    [Test]
    public async Task Range_InScalarContext_IsValueError()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));

        await Assert
            .That(ExpressionParser.Parse("=A1:A3", sheet).Evaluate(workbook).AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task UnknownFunction_IsNameError()
    {
        await Assert.That(Calc("=FOO(1)")).IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task EmptySum_IsZero()
    {
        await Assert.That(Calc("=SUM()") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task EmptyAverage_IsDivByZero()
    {
        await Assert.That(Calc("=AVERAGE()")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task NestedFunctions()
    {
        var (workbook, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));

        var result =
            ExpressionParser.Parse("=SUM(SUM(A1,A2),A3)", sheet).Evaluate(workbook).AsObject()
            as double?;

        await Assert.That(result).IsEqualTo(6.0);
    }

    [Test]
    public async Task ErrorFromFunction_PropagatesThroughArithmetic()
    {
        var (workbook, sheet) = Grid(("A1", 1));

        await Assert
            .That(ExpressionParser.Parse("=A1+FOO()", sheet).Evaluate(workbook).AsObject())
            .IsEqualTo(ErrorValue.Name);
    }

    // --- Syntax errors throw ParseException ---

    [Test]
    public async Task TrailingComma_IsOmittedArgument()
    {
        var (workbook, sheet) = Grid(("A1", 5));

        // A trailing/omitted argument is blank (ignored by SUM), like Excel — not a syntax error.
        await Assert
            .That(
                ExpressionParser.Parse("=SUM(A1,)", sheet).Evaluate(workbook).AsObject() as double?
            )
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task MissingComma_Throws()
    {
        var (_, sheet) = Grid();

        await Assert
            .That(() => ExpressionParser.Parse("=SUM(A1 A2)", sheet))
            .Throws<ParseException>();
    }

    [Test]
    public async Task UnclosedParenthesis_Throws()
    {
        var (_, sheet) = Grid();

        await Assert.That(() => ExpressionParser.Parse("=(1+2", sheet)).Throws<ParseException>();
    }

    [Test]
    public async Task DanglingOperator_Throws()
    {
        var (_, sheet) = Grid();

        await Assert.That(() => ExpressionParser.Parse("=*2", sheet)).Throws<ParseException>();
    }

    [Test]
    public async Task TrailingTokens_Throw()
    {
        var (_, sheet) = Grid();

        await Assert.That(() => ExpressionParser.Parse("=1 2", sheet)).Throws<ParseException>();
    }

    [Test]
    public async Task EmptyFormula_Throws()
    {
        var (_, sheet) = Grid();

        await Assert.That(() => ExpressionParser.Parse("=", sheet)).Throws<ParseException>();
    }

    // --- Recursion-depth guard (regression: a pathological formula must throw ParseException, never
    // overflow the stack — see Parser.MaxDepth) ---

    [Test]
    public async Task DeepNesting_WithinLimit_Parses()
    {
        var (_, sheet) = Grid();

        // 200 nested parentheses: well below MaxDepth (256), so this is a legitimate, if unusual, formula.
        var formula = "=" + new string('(', 200) + "1" + new string(')', 200);

        var expression = ExpressionParser.Parse(formula, sheet);

        await Assert.That(expression).IsTypeOf<NumberValue>();
    }

    [Test]
    public async Task DeepNesting_ExceedsLimit_ThrowsParseException()
    {
        var (_, sheet) = Grid();

        // 300 nested parentheses exceed MaxDepth; this must throw a catchable ParseException instead of
        // overflowing the stack.
        var formula = "=" + new string('(', 300) + "1" + new string(')', 300);

        await Assert.That(() => ExpressionParser.Parse(formula, sheet)).Throws<ParseException>();
    }

    [Test]
    public async Task DeepNesting_NestedFunctionCalls_ExceedsLimit_ThrowsParseException()
    {
        var (_, sheet) = Grid();

        // ABS(ABS(...(1)...)) 300 levels deep: nested function-call arguments recurse through
        // ParseExpression exactly like nested parentheses, so the same guard must catch it.
        var formula =
            "=" + string.Concat(Enumerable.Repeat("ABS(", 300)) + "1" + new string(')', 300);

        await Assert.That(() => ExpressionParser.Parse(formula, sheet)).Throws<ParseException>();
    }

    [Test]
    public async Task DeepNesting_ChainedUnaryMinus_ExceedsLimit_ThrowsParseException()
    {
        var (_, sheet) = Grid();

        // '----...-5' 300 levels deep: chained unary prefixes recurse through ParseExpression too.
        var formula = "=" + new string('-', 300) + "5";

        await Assert.That(() => ExpressionParser.Parse(formula, sheet)).Throws<ParseException>();
    }

    [Test]
    public async Task DeepNesting_ChainedQualifiedRangeEndpoints_ExceedsLimit_ThrowsParseException()
    {
        var (workbook, sheet) = Grid();
        workbook.Sheets.Add("S1");

        // Sheet1!A1:S1!A1:S1!A1:...: chained cross-sheet range endpoints recurse through
        // ParseQualifiedReference directly, WITHOUT going through ParseExpression — a separate cycle the
        // guard must also cover.
        var formula = "=Sheet1!A1" + string.Concat(Enumerable.Repeat(":S1!A1", 300));

        await Assert.That(() => ExpressionParser.Parse(formula, sheet)).Throws<ParseException>();
    }

    // --- Literal (non '=') entry mode ---

    [Test]
    public async Task Literal_Number()
    {
        await Assert.That(Calc("3") as double?).IsEqualTo(3.0);
    }

    [Test]
    public async Task Literal_Boolean()
    {
        await Assert.That(Calc("TRUE") as bool?).IsTrue();
    }

    [Test]
    public async Task Literal_Text()
    {
        await Assert.That(Calc("abc") as string).IsEqualTo("abc");
    }

    [Test]
    public async Task Literal_Empty_IsBlank()
    {
        await Assert.That(Calc("")).IsNull();
    }

    [Test]
    public async Task Literal_FormulaLikeTextIsNotComputed()
    {
        await Assert.That(Calc("1+1") as string).IsEqualTo("1+1");
    }
}
