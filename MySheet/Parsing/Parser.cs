using System.Globalization;
using MySheet.Expressions;

namespace MySheet.Parsing;

/// <summary>
/// A Pratt (top-down operator precedence) parser turning a token stream into an
/// <see cref="Expression"/> tree. Cell references are resolved against <c>sheetName</c>.
/// </summary>
internal sealed class Parser(List<Token> tokens, string sheetName)
{
    // Binding powers (higher binds tighter). Unary prefix binds tighter than '^' so that
    // '-2^2' parses as '(-2)^2' == 4, matching Excel.
    private const int ComparisonBindingPower = 10;
    private const int ConcatBindingPower = 15; // '&' binds below + - and above the comparators
    private const int AdditiveBindingPower = 20;
    private const int MultiplicativeBindingPower = 30;
    private const int PowerBindingPower = 40;
    private const int PercentBindingPower = 44; // postfix '%' binds above '^', below unary minus
    private const int PrefixBindingPower = 45;
    private const int RangeBindingPower = 50;

    private readonly record struct FunctionSpec(int MinArgs, int MaxArgs, Func<Expression[], Expression> Create);

    private static readonly Dictionary<string, FunctionSpec> Functions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUM"] = new(0, int.MaxValue, arguments => new Sum(arguments)),
            ["AVERAGE"] = new(0, int.MaxValue, arguments => new Average(arguments)),
            ["MIN"] = new(0, int.MaxValue, arguments => new Min(arguments)),
            ["MAX"] = new(0, int.MaxValue, arguments => new Max(arguments)),
            ["COUNT"] = new(0, int.MaxValue, arguments => new Count(arguments)),
            ["IF"] = new(2, 3, arguments => new If(arguments)),
            ["AND"] = new(1, int.MaxValue, arguments => new And(arguments)),
            ["OR"] = new(1, int.MaxValue, arguments => new Or(arguments)),
            ["NOT"] = new(1, 1, arguments => new Not(arguments)),
            ["IFERROR"] = new(2, 2, arguments => new IfError(arguments)),
            ["INT"] = new(1, 1, arguments => new Int(arguments)),
            ["ROUND"] = new(2, 2, arguments => new Round(arguments)),
            ["ROUNDUP"] = new(2, 2, arguments => new RoundUp(arguments)),
            ["ABS"] = new(1, 1, arguments => new Abs(arguments)),
            ["ISNUMBER"] = new(1, 1, arguments => new IsNumber(arguments)),
            ["ISBLANK"] = new(1, 1, arguments => new IsBlank(arguments)),
            ["IFNA"] = new(2, 2, arguments => new IfNa(arguments)),
            ["UPPER"] = new(1, 1, arguments => new Upper(arguments)),
            ["LOWER"] = new(1, 1, arguments => new Lower(arguments)),
            ["TRIM"] = new(1, 1, arguments => new Trim(arguments)),
            ["LEN"] = new(1, 1, arguments => new Len(arguments)),
            ["LEFT"] = new(1, 2, arguments => new Left(arguments)),
            ["MID"] = new(3, 3, arguments => new Mid(arguments)),
            ["VALUE"] = new(1, 1, arguments => new Value(arguments)),
            ["CONCAT"] = new(1, int.MaxValue, arguments => new Concat(arguments)),
            ["CONCATENATE"] = new(1, int.MaxValue, arguments => new Concatenate(arguments)),
            ["TEXTJOIN"] = new(3, int.MaxValue, arguments => new TextJoin(arguments)),
            ["COUNTA"] = new(1, int.MaxValue, arguments => new CountA(arguments)),
            ["COUNTBLANK"] = new(1, int.MaxValue, arguments => new CountBlank(arguments)),
            ["COUNTIF"] = new(2, 2, arguments => new CountIf(arguments)),
            ["COUNTIFS"] = new(2, int.MaxValue, arguments => new CountIfs(arguments)),
            ["SUMIF"] = new(2, 3, arguments => new SumIf(arguments)),
            ["SUMIFS"] = new(3, int.MaxValue, arguments => new SumIfs(arguments)),
            ["ROWS"] = new(1, 1, arguments => new Rows(arguments)),
            ["ROW"] = new(0, 1, arguments => new Row(arguments)),
            ["MATCH"] = new(2, 3, arguments => new Match(arguments)),
            ["INDEX"] = new(2, 3, arguments => new MySheet.Expressions.Index(arguments)),
            ["VLOOKUP"] = new(3, 4, arguments => new VLookup(arguments)),
            ["XLOOKUP"] = new(3, 6, arguments => new XLookup(arguments)),
            ["OFFSET"] = new(3, 5, arguments => new Offset(arguments)),
            ["LET"] = new(3, int.MaxValue, arguments => new Let(arguments)),
            ["TEXT"] = new(2, 2, arguments => new Text(arguments)),
            ["SHEET"] = new(0, 1, arguments => new SheetNumber(arguments)),
        };

    private int _index;

    private Token Current => tokens[_index];

    public Expression ParseFormula()
    {
        var expression = ParseExpression(0);

        if (Current.Type != TokenType.EndOfInput)
        {
            throw new ParseException($"Unexpected token '{Current.Text}'", Current.Position);
        }

        return expression;
    }

    private Expression ParseExpression(int rightBindingPower)
    {
        var left = ParsePrefix(Advance());

        while (rightBindingPower < LeftBindingPower(Current.Type))
        {
            left = ParseInfix(Advance(), left);
        }

        return left;
    }

    private Expression ParsePrefix(Token token)
    {
        switch (token.Type)
        {
            case TokenType.Number:
                return new NumberValue(double.Parse(token.Text, CultureInfo.InvariantCulture));

            case TokenType.String:
                return new StringValue(token.Text);

            case TokenType.Identifier:
                return ParseIdentifier(token);

            case TokenType.Minus:
                return new UnaryOperation(UnaryOperator.Negate, ParseExpression(PrefixBindingPower));

            case TokenType.Plus:
                return new UnaryOperation(UnaryOperator.Plus, ParseExpression(PrefixBindingPower));

            case TokenType.LParen:
                var inner = ParseExpression(0);
                Expect(TokenType.RParen);
                return inner;

            default:
                throw new ParseException($"Unexpected token '{token.Text}'", token.Position);
        }
    }

    private Expression ParseInfix(Token op, Expression left)
    {
        if (op.Type == TokenType.Colon)
        {
            return ParseRange(op, left);
        }

        // '%' is postfix: it divides the value to its left by 100, with no right operand.
        if (op.Type == TokenType.Percent)
        {
            return new UnaryOperation(UnaryOperator.Percent, left);
        }

        var bindingPower = LeftBindingPower(op.Type);
        var rightAssociative = op.Type == TokenType.Caret;
        var right = ParseExpression(rightAssociative ? bindingPower - 1 : bindingPower);

        return new BinaryOperation(ToBinaryOperator(op.Type), left, right);
    }

    private Expression ParseRange(Token colon, Expression left)
    {
        var right = ParseExpression(RangeBindingPower);

        if (left is CellReference start && right is CellReference end)
        {
            // The range lives on the start cell's sheet (e.g. Sheet2!A1:B2 is all on Sheet2).
            return new RangeReference(start.Id, end.Id, start.SheetName);
        }

        throw new ParseException("The ':' range operator requires cell references", colon.Position);
    }

    private Expression ParseIdentifier(Token token)
    {
        // A name before '!' is a sheet qualifier: Sheet2!A1, 'My Sheet'!A1.
        if (Current.Type == TokenType.Bang)
        {
            return ParseQualifiedReference(token.Text);
        }

        if (Current.Type == TokenType.LParen)
        {
            return ParseFunctionCall(token);
        }

        if (IsBoolean(token.Text, out var boolean))
        {
            return new BooleanValue(boolean);
        }

        if (IsCellReference(token.Text))
        {
            return new CellReference(NormalizeCellId(token.Text), sheetName);
        }

        // A bare name: a LET-bound name resolved at evaluation time (#NAME? if unbound).
        return new NameReference(token.Text);
    }

    private Expression ParseQualifiedReference(string sheet)
    {
        Expect(TokenType.Bang);
        var token = Advance();

        if (token.Type != TokenType.Identifier || !IsCellReference(token.Text))
        {
            throw new ParseException("Expected a cell reference after '!'", token.Position);
        }

        return new CellReference(NormalizeCellId(token.Text), sheet);
    }

    // Strips absolute markers ('$') and upper-cases, e.g. $A$1 -> A1. The reference identifies the same
    // cell regardless of '$'; absolute/relative only matters for Excel copy/fill, which we do not do.
    private static string NormalizeCellId(string text) => StripDollars(text).ToUpperInvariant();

    private static string StripDollars(string text) => text.Contains('$') ? text.Replace("$", string.Empty) : text;

    private Expression ParseFunctionCall(Token name)
    {
        Expect(TokenType.LParen);

        var arguments = new List<Expression>();

        if (Current.Type != TokenType.RParen)
        {
            arguments.Add(ParseArgument());

            while (Current.Type == TokenType.Comma)
            {
                Advance();
                arguments.Add(ParseArgument());
            }
        }

        Expect(TokenType.RParen);

        var functionName = NormalizeFunctionName(name.Text);

        // Built-in: typed record with parse-time arity validation (a wrong count throws, like Excel
        // rejecting it at entry). Otherwise a generic call resolved at runtime against the workbook's
        // custom-function registry (#NAME? if never registered).
        if (!Functions.TryGetValue(functionName, out var spec))
        {
            return new FunctionCall(functionName, arguments.ToArray());
        }

        if (arguments.Count < spec.MinArgs || arguments.Count > spec.MaxArgs)
        {
            throw new ParseException(
                $"Function '{functionName}' does not accept {arguments.Count} argument(s)", name.Position);
        }

        return spec.Create(arguments.ToArray());
    }

    private Expression ParseArgument()
    {
        // An omitted argument (e.g. XLOOKUP(a,b,c,,2) or a trailing comma) is treated as blank.
        if (Current.Type is TokenType.Comma or TokenType.RParen)
        {
            return BlankValue.Instance;
        }

        return ParseExpression(0);
    }

    // Excel stores newer functions with an "_xlfn." prefix; normalize it (and the bare "XLFN.") away.
    private static string NormalizeFunctionName(string name)
    {
        if (name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase))
        {
            return name["_xlfn.".Length..];
        }

        if (name.StartsWith("xlfn.", StringComparison.OrdinalIgnoreCase))
        {
            return name["xlfn.".Length..];
        }

        return name;
    }

    private Token Advance()
    {
        var token = tokens[_index];

        if (token.Type != TokenType.EndOfInput)
        {
            _index++;
        }

        return token;
    }

    private void Expect(TokenType type)
    {
        if (Current.Type != type)
        {
            throw new ParseException($"Expected {type} but found '{Current.Text}'", Current.Position);
        }

        Advance();
    }

    private static int LeftBindingPower(TokenType type) => type switch
    {
        TokenType.Equal or TokenType.NotEqual or TokenType.Less or TokenType.Greater
            or TokenType.LessEqual or TokenType.GreaterEqual => ComparisonBindingPower,
        TokenType.Ampersand => ConcatBindingPower,
        TokenType.Plus or TokenType.Minus => AdditiveBindingPower,
        TokenType.Star or TokenType.Slash => MultiplicativeBindingPower,
        TokenType.Caret => PowerBindingPower,
        TokenType.Percent => PercentBindingPower,
        TokenType.Colon => RangeBindingPower,
        _ => 0,
    };

    private static BinaryOperator ToBinaryOperator(TokenType type) => type switch
    {
        TokenType.Plus => BinaryOperator.Add,
        TokenType.Minus => BinaryOperator.Subtract,
        TokenType.Star => BinaryOperator.Multiply,
        TokenType.Slash => BinaryOperator.Divide,
        TokenType.Caret => BinaryOperator.Power,
        TokenType.Equal => BinaryOperator.Equal,
        TokenType.NotEqual => BinaryOperator.NotEqual,
        TokenType.Less => BinaryOperator.LessThan,
        TokenType.Greater => BinaryOperator.GreaterThan,
        TokenType.LessEqual => BinaryOperator.LessThanOrEqual,
        TokenType.GreaterEqual => BinaryOperator.GreaterThanOrEqual,
        TokenType.Ampersand => BinaryOperator.Concat,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null),
    };

    private static bool IsBoolean(string text, out bool value)
    {
        if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return true;
        }

        if (string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return true;
        }

        value = false;
        return false;
    }

    private static bool IsCellReference(string text)
    {
        text = StripDollars(text);

        var letters = 0;
        while (letters < text.Length && char.IsLetter(text[letters]))
        {
            letters++;
        }

        if (letters == 0 || letters == text.Length)
        {
            return false;
        }

        for (var i = letters; i < text.Length; i++)
        {
            if (!char.IsDigit(text[i]))
            {
                return false;
            }
        }

        return true;
    }
}
