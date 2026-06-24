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
    private const int AdditiveBindingPower = 20;
    private const int MultiplicativeBindingPower = 30;
    private const int PowerBindingPower = 40;
    private const int PrefixBindingPower = 45;
    private const int RangeBindingPower = 50;

    private static readonly Dictionary<string, Func<Expression[], Expression>> Functions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUM"] = arguments => new Sum(arguments),
            ["AVERAGE"] = arguments => new Average(arguments),
            ["MIN"] = arguments => new Min(arguments),
            ["MAX"] = arguments => new Max(arguments),
            ["COUNT"] = arguments => new Count(arguments),
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
            return new RangeReference(start.Id, end.Id, sheetName);
        }

        throw new ParseException("The ':' range operator requires cell references", colon.Position);
    }

    private Expression ParseIdentifier(Token token)
    {
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
            return new CellReference(token.Text.ToUpperInvariant(), sheetName);
        }

        // Unknown name: a semantic error, surfaced as a node rather than thrown.
        return ErrorValue.Name;
    }

    private Expression ParseFunctionCall(Token name)
    {
        Expect(TokenType.LParen);

        var arguments = new List<Expression>();

        if (Current.Type != TokenType.RParen)
        {
            arguments.Add(ParseExpression(0));

            while (Current.Type == TokenType.Comma)
            {
                Advance();
                arguments.Add(ParseExpression(0));
            }
        }

        Expect(TokenType.RParen);

        return Functions.TryGetValue(name.Text, out var factory)
            ? factory(arguments.ToArray())
            : ErrorValue.Name;
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
        TokenType.Plus or TokenType.Minus => AdditiveBindingPower,
        TokenType.Star or TokenType.Slash => MultiplicativeBindingPower,
        TokenType.Caret => PowerBindingPower,
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
