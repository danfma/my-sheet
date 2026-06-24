using System.Text;

namespace MySheet.Parsing;

/// <summary>
/// Splits a formula body (after the leading '=' has been stripped) into a flat list of tokens
/// terminated by an <see cref="TokenType.EndOfInput"/> token. Whitespace is discarded.
/// </summary>
internal sealed class Tokenizer(string text)
{
    private int _position;

    public static List<Token> Tokenize(string text)
    {
        var tokenizer = new Tokenizer(text);
        var tokens = new List<Token>();

        Token token;
        do
        {
            token = tokenizer.NextToken();
            tokens.Add(token);
        }
        while (token.Type != TokenType.EndOfInput);

        return tokens;
    }

    private Token NextToken()
    {
        SkipWhitespace();

        if (_position >= text.Length)
        {
            return new Token(TokenType.EndOfInput, string.Empty, _position);
        }

        var start = _position;
        var c = text[_position];

        if (char.IsDigit(c) || c == '.')
        {
            return ReadNumber(start);
        }

        if (char.IsLetter(c) || c is '_' or '$')
        {
            return ReadIdentifier(start);
        }

        if (c == '"')
        {
            return ReadString(start);
        }

        if (c == '\'')
        {
            return ReadQuotedName(start);
        }

        return ReadOperator(start);
    }

    private void SkipWhitespace()
    {
        while (_position < text.Length && char.IsWhiteSpace(text[_position]))
        {
            _position++;
        }
    }

    private Token ReadNumber(int start)
    {
        while (_position < text.Length && (char.IsDigit(text[_position]) || text[_position] == '.'))
        {
            _position++;
        }

        // Scientific notation: e/E, an optional sign, then at least one digit. Otherwise leave the
        // 'e' for the next token (so a trailing 'E' does not get swallowed into the number).
        if (_position < text.Length && (text[_position] is 'e' or 'E'))
        {
            var exponentStart = _position;
            _position++;

            if (_position < text.Length && (text[_position] is '+' or '-'))
            {
                _position++;
            }

            if (_position < text.Length && char.IsDigit(text[_position]))
            {
                while (_position < text.Length && char.IsDigit(text[_position]))
                {
                    _position++;
                }
            }
            else
            {
                _position = exponentStart;
            }
        }

        return new Token(TokenType.Number, text[start.._position], start);
    }

    private Token ReadIdentifier(int start)
    {
        // Names may contain '_' (A_HIDE), '.' (XLFN.XLOOKUP) and '$' (absolute refs like $A$1); cell-ref
        // classification in the parser strips '$' and requires the strict [A-Za-z]+[0-9]+ shape.
        while (_position < text.Length && (char.IsLetterOrDigit(text[_position]) || text[_position] is '_' or '.' or '$'))
        {
            _position++;
        }

        return new Token(TokenType.Identifier, text[start.._position], start);
    }

    private Token ReadString(int start)
    {
        _position++; // opening quote
        var builder = new StringBuilder();

        while (_position < text.Length)
        {
            var c = text[_position];

            if (c == '"')
            {
                // A doubled quote ("") is an escaped quote inside the string.
                if (_position + 1 < text.Length && text[_position + 1] == '"')
                {
                    builder.Append('"');
                    _position += 2;
                    continue;
                }

                _position++; // closing quote
                return new Token(TokenType.String, builder.ToString(), start);
            }

            builder.Append(c);
            _position++;
        }

        throw new ParseException("Unterminated string literal", start);
    }

    // A sheet name in single quotes (allows spaces/specials), e.g. 'My Sheet'!A1. '' is an escaped quote.
    private Token ReadQuotedName(int start)
    {
        _position++; // opening quote
        var builder = new StringBuilder();

        while (_position < text.Length)
        {
            var c = text[_position];

            if (c == '\'')
            {
                if (_position + 1 < text.Length && text[_position + 1] == '\'')
                {
                    builder.Append('\'');
                    _position += 2;
                    continue;
                }

                _position++; // closing quote
                return new Token(TokenType.Identifier, builder.ToString(), start);
            }

            builder.Append(c);
            _position++;
        }

        throw new ParseException("Unterminated quoted name", start);
    }

    private Token ReadOperator(int start)
    {
        var c = text[_position];

        switch (c)
        {
            case '+': _position++; return new Token(TokenType.Plus, "+", start);
            case '-': _position++; return new Token(TokenType.Minus, "-", start);
            case '*': _position++; return new Token(TokenType.Star, "*", start);
            case '/': _position++; return new Token(TokenType.Slash, "/", start);
            case '^': _position++; return new Token(TokenType.Caret, "^", start);
            case ',': _position++; return new Token(TokenType.Comma, ",", start);
            case ':': _position++; return new Token(TokenType.Colon, ":", start);
            case '(': _position++; return new Token(TokenType.LParen, "(", start);
            case ')': _position++; return new Token(TokenType.RParen, ")", start);
            case '=': _position++; return new Token(TokenType.Equal, "=", start);
            case '!': _position++; return new Token(TokenType.Bang, "!", start);
            case '&': _position++; return new Token(TokenType.Ampersand, "&", start);

            case '<':
                _position++;
                if (Match('>')) return new Token(TokenType.NotEqual, "<>", start);
                if (Match('=')) return new Token(TokenType.LessEqual, "<=", start);
                return new Token(TokenType.Less, "<", start);

            case '>':
                _position++;
                if (Match('=')) return new Token(TokenType.GreaterEqual, ">=", start);
                return new Token(TokenType.Greater, ">", start);

            default:
                throw new ParseException($"Unexpected character '{c}'", start);
        }
    }

    private bool Match(char expected)
    {
        if (_position < text.Length && text[_position] == expected)
        {
            _position++;
            return true;
        }

        return false;
    }
}
