using MySheet.Parsing;

namespace MySheet.Tests.Parsing;

public class TokenizerTests
{
    private static string Shape(string input) =>
        string.Join(" ", Tokenizer.Tokenize(input).Select(t => t.Type));

    private static Token Single(string input, TokenType type) =>
        Tokenizer.Tokenize(input).Single(t => t.Type == type);

    [Test]
    public async Task Tokenizes_FunctionCall()
    {
        await Assert.That(Shape("SUM(A1,A2)"))
            .IsEqualTo("Identifier LParen Identifier Comma Identifier RParen EndOfInput");
    }

    [Test]
    public async Task Tokenizes_Range()
    {
        await Assert.That(Shape("A1:B2"))
            .IsEqualTo("Identifier Colon Identifier EndOfInput");
    }

    [Test]
    public async Task Tokenizes_Arithmetic_WithUnaryMinus()
    {
        await Assert.That(Shape("-3 + 4 * 2"))
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
        await Assert.That(Shape("SUM (A1)"))
            .IsEqualTo("Identifier LParen Identifier RParen EndOfInput");
    }

    [Test]
    public async Task InvalidCharacter_Throws()
    {
        await Assert.That(() => Tokenizer.Tokenize("1 # 2")).Throws<ParseException>();
    }
}
