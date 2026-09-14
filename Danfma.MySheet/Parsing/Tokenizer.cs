using System.Text;

namespace Danfma.MySheet.Parsing;

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
        // Presized: one token per ~3 chars is a safe upper-bound shape for real formulas, so the
        // list virtually never regrows.
        var tokens = new List<Token>(text.Length / 3 + 4);

        Token token;
        do
        {
            token = tokenizer.NextToken();
            tokens.Add(token);
        } while (token.Type != TokenType.EndOfInput);

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

        // After the '\'' case on purpose: the quoted external form '[1]Data'!A1 must stay a quoted name,
        // so the bracket reader only ever sees a '[' that is not inside quotes.
        if (c == '[')
        {
            return ReadBracketedSpecifier(start);
        }

        if (c == '#')
        {
            return ReadErrorLiteral(start);
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
        while (
            _position < text.Length
            && (char.IsLetterOrDigit(text[_position]) || text[_position] is '_' or '.' or '$')
        )
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

        throw new ParseException(
            ParseErrorKind.UnterminatedString,
            "Unterminated string literal",
            start,
            text[start..]
        );
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

        throw new ParseException(
            ParseErrorKind.UnterminatedQuotedName,
            "Unterminated quoted name",
            start,
            text[start..]
        );
    }

    // The whole `[...]` suffix of a structured (table) reference, e.g. the `[Valor]` of Tabela1[Valor] or
    // the `[[#Data],[% Comissao]]` of the composite form. Unconditional and context-free like every other
    // reader here (no "previous token" gate): a `[` anywhere becomes one token, and the parser is what
    // rejects the out-of-scope shapes (`[Valor]` alone, `[1]Sheet1!A1`) with a message naming them.
    //
    // Deliberate divergence from ReadQuotedName: Text is the RAW slice, brackets included and undecoded.
    // Decoding (the `'` escape) is per item and happens after the split in StructuredReferenceSyntax, and
    // keeping the brackets means ParseException.Token and the parser's generic messages print `[Valor]`,
    // the text as it appeared in the formula. The scan itself is StructuredReferenceSyntax's — the one
    // balanced-bracket scanner shared with the item splitter — so a payload this reader accepts is one
    // the splitter can split; a raw newline inside the brackets is payload, as Excel stores it.
    private Token ReadBracketedSpecifier(int start)
    {
        if (!StructuredReferenceSyntax.TryFindClosingBracket(text, _position, out var close))
        {
            throw new ParseException(
                ParseErrorKind.UnterminatedBracketedReference,
                "Unterminated bracketed reference",
                start,
                text[start..]
            );
        }

        var token = new Token(TokenType.BracketedSpecifier, text[start..(close + 1)], start);
        _position = close + 1;
        return token;
    }

    // The 7 classic Excel error literals, in their canonical (upper-case) spelling — the same set
    // Danfma.MySheet.Error models (#CALC! excluded: Aspose.Cells 26.7.0/26.6.0 both refuse to PARSE it as
    // formula-text syntax, "Invalid '#'", even though it is a real error CODE MySheet can hold — Excel
    // only ever PRODUCES it). None is a prefix of another, so a simple ordered scan is unambiguous.
    private static readonly string[] ErrorLiterals =
    [
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
    ];

    // Item 43 (sweep 31-35-43): reads an error literal case-insensitively, returning the CANONICAL
    // spelling as the token's Text so the parser can hand it straight to Error.FromDisplay with no
    // second normalization step. A '#' that does not open one of the 7 known spellings is the same
    // syntax error it always was (Aspose.Cells rejects it too, e.g. #SPILL!/#CALC!/#GETTING_DATA:
    // "Invalid '#'" — measured 2026-09-14).
    private Token ReadErrorLiteral(int start)
    {
        foreach (var literal in ErrorLiterals)
        {
            if (
                start + literal.Length <= text.Length
                && string.Compare(
                    text,
                    start,
                    literal,
                    0,
                    literal.Length,
                    StringComparison.OrdinalIgnoreCase
                ) == 0
            )
            {
                _position = start + literal.Length;
                return new Token(TokenType.Error, literal, start);
            }
        }

        throw new ParseException(
            ParseErrorKind.UnexpectedCharacter,
            "Unexpected character '#'",
            start,
            "#"
        );
    }

    private Token ReadOperator(int start)
    {
        var c = text[_position];

        switch (c)
        {
            case '+':
                _position++;
                return new Token(TokenType.Plus, "+", start);
            case '-':
                _position++;
                return new Token(TokenType.Minus, "-", start);
            case '*':
                _position++;
                return new Token(TokenType.Star, "*", start);
            case '/':
                _position++;
                return new Token(TokenType.Slash, "/", start);
            case '^':
                _position++;
                return new Token(TokenType.Caret, "^", start);
            case ',':
                _position++;
                return new Token(TokenType.Comma, ",", start);
            case ':':
                _position++;
                return new Token(TokenType.Colon, ":", start);
            case '(':
                _position++;
                return new Token(TokenType.LParen, "(", start);
            case ')':
                _position++;
                return new Token(TokenType.RParen, ")", start);
            case '=':
                _position++;
                return new Token(TokenType.Equal, "=", start);
            case '!':
                _position++;
                return new Token(TokenType.Bang, "!", start);
            case '&':
                _position++;
                return new Token(TokenType.Ampersand, "&", start);
            case '%':
                _position++;
                return new Token(TokenType.Percent, "%", start);

            case '<':
                _position++;
                if (Match('>'))
                    return new Token(TokenType.NotEqual, "<>", start);
                if (Match('='))
                    return new Token(TokenType.LessEqual, "<=", start);
                return new Token(TokenType.Less, "<", start);

            case '>':
                _position++;
                if (Match('='))
                    return new Token(TokenType.GreaterEqual, ">=", start);
                return new Token(TokenType.Greater, ">", start);

            default:
                throw new ParseException(
                    ParseErrorKind.UnexpectedCharacter,
                    $"Unexpected character '{c}'",
                    start,
                    c.ToString()
                );
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
