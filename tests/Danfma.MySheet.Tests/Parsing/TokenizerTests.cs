using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class TokenizerTests
{
    private static string Shape(string input) =>
        string.Join(" ", Tokenizer.Tokenize(input).Select(t => t.Type));

    private static Token Single(string input, TokenType type) =>
        Tokenizer.Tokenize(input).Single(t => t.Type == type);

    [Test]
    public async Task Tokenizes_FunctionCall()
    {
        await Assert
            .That(Shape("SUM(A1,A2)"))
            .IsEqualTo("Identifier LParen Identifier Comma Identifier RParen EndOfInput");
    }

    [Test]
    public async Task Tokenizes_Range()
    {
        await Assert.That(Shape("A1:B2")).IsEqualTo("Identifier Colon Identifier EndOfInput");
    }

    [Test]
    public async Task Tokenizes_Arithmetic_WithUnaryMinus()
    {
        await Assert
            .That(Shape("-3 + 4 * 2"))
            .IsEqualTo("Minus Number Plus Number Star Number EndOfInput");
    }

    [Test]
    public async Task Tokenizes_ComparisonOperators()
    {
        await Assert.That(Shape("1 <= 2")).IsEqualTo("Number LessEqual Number EndOfInput");
        await Assert.That(Shape("1<>2")).IsEqualTo("Number NotEqual Number EndOfInput");
        await Assert.That(Shape("1>=2")).IsEqualTo("Number GreaterEqual Number EndOfInput");
    }

    [Test]
    public async Task ScientificNotation_IsASingleNumberToken()
    {
        await Assert.That(Shape("1E2")).IsEqualTo("Number EndOfInput");
        await Assert.That(Single("1E2", TokenType.Number).Text).IsEqualTo("1E2");
        await Assert.That(Single("1.5E-3", TokenType.Number).Text).IsEqualTo("1.5E-3");
    }

    [Test]
    public async Task CellReference_LikeE2_IsAnIdentifier()
    {
        await Assert.That(Shape("E2")).IsEqualTo("Identifier EndOfInput");
    }

    [Test]
    public async Task String_UnescapesDoubledQuotes()
    {
        var token = Single("\"a\"\"b\"", TokenType.String);

        await Assert.That(token.Text).IsEqualTo("a\"b");
    }

    [Test]
    public async Task Whitespace_BeforeParen_IsDiscarded()
    {
        await Assert
            .That(Shape("SUM (A1)"))
            .IsEqualTo("Identifier LParen Identifier RParen EndOfInput");
    }

    [Test]
    public async Task InvalidCharacter_Throws()
    {
        await Assert.That(() => Tokenizer.Tokenize("1 # 2")).Throws<ParseException>();
    }

    [Test]
    public async Task StructuredReference_IsOneToken()
    {
        await Assert
            .That(Shape("Tabela1[Valor]"))
            .IsEqualTo("Identifier BracketedSpecifier EndOfInput");

        // `#`, `,`, `(`, `)` and a space inside the brackets are payload, not operators — the same
        // three-token shape as the plain column.
        await Assert
            .That(Shape("Tabela1[[#Data],[% Comissao]]"))
            .IsEqualTo("Identifier BracketedSpecifier EndOfInput");
        await Assert
            .That(Shape("Tabela1[Total (USD)]"))
            .IsEqualTo("Identifier BracketedSpecifier EndOfInput");
    }

    [Test]
    public async Task BracketedSpecifier_KeepsTheRawTextIncludingBrackets()
    {
        var token = Single("Tabela1[Total (USD)]", TokenType.BracketedSpecifier);

        await Assert.That(token.Text).IsEqualTo("[Total (USD)]");
        await Assert.That(token.Position).IsEqualTo(7);
    }

    [Test]
    public async Task UnterminatedBracket_Throws()
    {
        var exception = Assert.Throws<ParseException>(() => Tokenizer.Tokenize("Tabela1[Valor"));

        await Assert.That(exception.Kind).IsEqualTo(ParseErrorKind.UnterminatedBracketedReference);
        await Assert.That(exception.Position).IsEqualTo(7);
        // Same convention as ReadString/ReadQuotedName: the unterminated payload from the opener on.
        await Assert.That(exception.Token).IsEqualTo("[Valor");
    }

    [Test]
    public async Task StructuredReference_TerminatesAtTheMatchingBracket()
    {
        // The first specifier closes at the `]` that balances its `[` (index 13), not at the first
        // `]` — so `Table[a]:Table2[b]` reaches the parser as two references around a Colon, where
        // TryEndpoint (which accepts only CellReference/NameReference/NumberValue) lets it fall
        // through to a DynamicRange spanning two table columns.
        await Assert
            .That(Shape("T[[#Data],[a]]:T2[b]"))
            .IsEqualTo(
                "Identifier BracketedSpecifier Colon Identifier BracketedSpecifier EndOfInput"
            );
    }

    [Test]
    public async Task BracketedSpecifier_KeepsARawNewline()
    {
        // Excel stores a column named with a line break as a real `\n` inside the cell's formula,
        // so the reader must not stop at a newline: it is payload.
        var tokens = Tokenizer.Tokenize("SUM(Tabela1[Line\nBreak])");

        await Assert
            .That(string.Join(" ", tokens.Select(t => t.Type)))
            .IsEqualTo("Identifier LParen Identifier BracketedSpecifier RParen EndOfInput");
        await Assert
            .That(tokens.Single(t => t.Type == TokenType.BracketedSpecifier).Text)
            .IsEqualTo("[Line\nBreak]");
    }

    [Test]
    [Arguments("Tabela1[@Valor]", "[@Valor]")]
    [Arguments("Tabela1[[#This Row],[Valor]]", "[[#This Row],[Valor]]")]
    public async Task ThisRow_BothSpellings_LexToOneBracketedSpecifier(
        string input,
        string expected
    )
    {
        // `[@Valor]` is what a user types; `[[#This Row],[Valor]]` is what a saved file carries. Both
        // are one raw token here; the parser decides what to do with them.
        await Assert.That(Shape(input)).IsEqualTo("Identifier BracketedSpecifier EndOfInput");
        await Assert.That(Single(input, TokenType.BracketedSpecifier).Text).IsEqualTo(expected);
    }

    // Item 43 (sweep 31-35-43): the 7 classic error literals Aspose.Cells 26.7.0/26.6.0 accept as
    // FORMULA-TEXT syntax (PLAIN and CSE alike, measured 2026-09-14). Each is one token, case-insensitive,
    // and the token's Text is always the CANONICAL (upper-case) spelling regardless of input case, so it
    // feeds `Error.FromDisplay` directly without a second normalization step downstream.
    [Test]
    [Arguments("#NULL!")]
    [Arguments("#DIV/0!")]
    [Arguments("#VALUE!")]
    [Arguments("#REF!")]
    [Arguments("#NAME?")]
    [Arguments("#NUM!")]
    [Arguments("#N/A")]
    public async Task ErrorLiteral_IsOneToken(string literal)
    {
        await Assert.That(Shape(literal)).IsEqualTo("Error EndOfInput");
        await Assert.That(Single(literal, TokenType.Error).Text).IsEqualTo(literal);
    }

    [Test]
    [Arguments("#ref!", "#REF!")]
    [Arguments("#Ref!", "#REF!")]
    [Arguments("#n/a", "#N/A")]
    [Arguments("#N/a", "#N/A")]
    [Arguments("#div/0!", "#DIV/0!")]
    [Arguments("#Div/0!", "#DIV/0!")]
    [Arguments("#name?", "#NAME?")]
    [Arguments("#Name?", "#NAME?")]
    public async Task ErrorLiteral_IsCaseInsensitive_AndCanonicalizes(
        string input,
        string canonical
    )
    {
        await Assert.That(Single(input, TokenType.Error).Text).IsEqualTo(canonical);
    }

    // Aspose.Cells 26.7.0/26.6.0 reject all three as formula-text SYNTAX ("Invalid '#'"), measured
    // 2026-09-14 — even #CALC!, which IS a real Excel error code MySheet already models
    // (ErrorValue.Calculation, Excel's empty-array result): Excel/Aspose can PRODUCE it, never accepts
    // TYPING it. So the tokenizer stays narrow to the 7 literals above; anything else starting with '#'
    // keeps throwing exactly as it did before this item (UnexpectedCharacter).
    [Test]
    [Arguments("#GETTING_DATA")]
    [Arguments("#SPILL!")]
    [Arguments("#CALC!")]
    [Arguments("#BOGUS!")]
    public async Task UnrecognizedErrorSpelling_StillThrows(string input)
    {
        var exception = Assert.Throws<ParseException>(() => Tokenizer.Tokenize(input));

        await Assert.That(exception.Kind).IsEqualTo(ParseErrorKind.UnexpectedCharacter);
        await Assert.That(exception.Token).IsEqualTo("#");
    }

    [Test]
    public async Task ErrorLiteralNearMiss_KeepsTheBaselineExceptionContract()
    {
        var exception = Assert.Throws<ParseException>(() => Tokenizer.Tokenize("#N/A!"));

        await Assert.That(exception.Kind).IsEqualTo(ParseErrorKind.UnexpectedCharacter);
        await Assert.That(exception.Token).IsEqualTo("#");
        await Assert.That(exception.Position).IsEqualTo(0);
    }

    [Test]
    public async Task QuotedExternalName_IsStillAQuotedName_NotABracket()
    {
        // `'` dispatches before `[`, so the quoted external form is untouched by the bracket reader...
        await Assert.That(Shape("'[1]Data'!A1")).IsEqualTo("Identifier Bang Identifier EndOfInput");
        await Assert.That(Tokenizer.Tokenize("'[1]Data'!A1")[0].Text).IsEqualTo("[1]Data");

        // ...while the unquoted external form now lexes as a bracketed specifier the parser rejects.
        await Assert
            .That(Shape("[1]Sheet1!A1"))
            .IsEqualTo("BracketedSpecifier Identifier Bang Identifier EndOfInput");
    }
}
