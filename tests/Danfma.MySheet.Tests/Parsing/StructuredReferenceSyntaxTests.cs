using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Phase 4 T1 (foundation): the pieces of the structured-reference syntax that every other Phase 4/5 task
/// builds on — the <see cref="TokenType.BracketedSpecifier"/> token member, the three
/// <see cref="ParseErrorKind"/> members, and the balanced-bracket scanner. The tokenizer dispatch to the
/// scanner (Phase 4 T3) and the item grammar over its result (Phase 4 T2) are pinned in their own files.
/// </summary>
public class StructuredReferenceSyntaxTests
{
    // === Item 1: the token member ========================================================================

    // TokenizerTests' golden strings join `t.Type` NAMES, so the member's name is the wire that matters,
    // and nothing in the engine uses TokenType ordinally (re-grepped at implementation time).
    [Test]
    public async Task BracketedSpecifier_IsATokenType_AndPrintsItsName()
    {
        var token = new Token(TokenType.BracketedSpecifier, "[Valor]", 7);

        await Assert.That(token.Type.ToString()).IsEqualTo("BracketedSpecifier");
        await Assert.That(token.Text).IsEqualTo("[Valor]");
    }

    // === Item 2: the three error kinds ===================================================================

    // ParseErrorKind is PUBLIC, so the three members are APPENDED after NestingTooDeep, in this order.
    [Test]
    public async Task StructuredReferenceKinds_AreAppendedAfterNestingTooDeep_InOrder()
    {
        var names = Enum.GetNames<ParseErrorKind>();
        var afterNestingTooDeep = names[(Array.IndexOf(names, "NestingTooDeep") + 1)..];

        await Assert
            .That(afterNestingTooDeep)
            .IsEquivalentTo([
                "UnterminatedBracketedReference",
                "InvalidStructuredReference",
                "UnsupportedStructuredReference",
            ]);
    }

    [Test]
    [Arguments(ParseErrorKind.UnterminatedBracketedReference)]
    [Arguments(ParseErrorKind.InvalidStructuredReference)]
    [Arguments(ParseErrorKind.UnsupportedStructuredReference)]
    public async Task StructuredReferenceKinds_TravelThroughParseException(ParseErrorKind kind)
    {
        var exception = new ParseException(kind, "message", 8, "[Valor");

        await Assert.That(exception.Kind).IsEqualTo(kind);
        await Assert.That(exception.Position).IsEqualTo(8);
        await Assert.That(exception.Token).IsEqualTo("[Valor");
    }

    // === Item 5 (scanner half): TryFindClosingBracket ===================================================

    // The one balanced-bracket scanner shared by the tokenizer's reader, the item splitter and the per-item
    // bracket check, so the tokenizer can never accept a payload the splitter mis-splits. It honours
    // Excel's `'` escape (the escape and its escapee are skipped as a pair) and nothing else — no decoding,
    // no trimming, no whitespace rule.
    [Test]
    [Arguments("Tabela1[Valor]", 7, 13)]
    [Arguments("Tabela1[Valor]+1", 7, 13)] // the FIRST balanced close, not the last ']' in the text
    [Arguments("T[#All]", 1, 6)]
    [Arguments("T[[#Headers],[#Data],[Valor]]", 1, 28)] // nested item list: the outermost close
    [Arguments("T[a']b]", 1, 6)] // '] is an escaped close bracket
    [Arguments("T[x'[y]", 1, 6)] // '[ is an escaped open bracket, so depth never reaches 2
    [Arguments("T[a''b]", 1, 6)] // '' is an escaped apostrophe: the first skips the second
    [Arguments("T[a'#b]", 1, 6)] // '# and '@ are ordinary to the scanner, still skipped as a pair
    [Arguments("T[ Padded ]", 1, 10)] // whitespace is payload, not a terminator
    public async Task TryFindClosingBracket_FindsTheBalancedClose(string text, int open, int close)
    {
        var ok = StructuredReferenceSyntax.TryFindClosingBracket(text, open, out var found);

        await Assert.That(ok).IsTrue();
        await Assert.That(found).IsEqualTo(close);
    }

    // Phase 6 addendum: Aspose stores a column named with a line break as a REAL '\n' inside the cell's
    // <f> (SUM(Tabela1[Line\nBreak])), so a scanner that stopped at a newline would degrade every such
    // formula. The tokenizer-level twin of this pin (Tokenize(...) yielding one BracketedSpecifier) is T3's.
    [Test]
    public async Task TryFindClosingBracket_AcceptsARawNewlineInsideTheBrackets()
    {
        const string text = "SUM(Tabela1[Line\nBreak])";
        var open = text.IndexOf('[');

        var ok = StructuredReferenceSyntax.TryFindClosingBracket(text, open, out var close);

        await Assert.That(ok).IsTrue();
        await Assert.That(text[close]).IsEqualTo(']');
        await Assert.That(text[(open + 1)..close]).IsEqualTo("Line\nBreak");
    }

    [Test]
    [Arguments("T[Valor", 1)] // unterminated
    [Arguments("T[[Valor]", 1)] // unbalanced: the inner item closes, the outer never does
    [Arguments("T[abc'", 1)] // a trailing escape with nothing to escape runs off the end
    [Arguments("T[abc']", 1)] // ...and so does one that escapes the only ']'
    public async Task TryFindClosingBracket_RunsOffTheEnd_IsFalse(string text, int open)
    {
        var ok = StructuredReferenceSyntax.TryFindClosingBracket(text, open, out var close);

        await Assert.That(ok).IsFalse();
        await Assert.That(close).IsEqualTo(-1);
    }

    // The item splitter calls this on each trimmed item with openIndex 0 and treats false as "not a
    // bracketed item", so a start that is not '[' (or is out of range) answers false rather than throwing.
    [Test]
    [Arguments("Valor", 0)]
    [Arguments("T[Valor]", 0)]
    [Arguments("T[Valor]", 8)]
    [Arguments("", 0)]
    public async Task TryFindClosingBracket_NotAtAnOpenBracket_IsFalse(string text, int open)
    {
        var ok = StructuredReferenceSyntax.TryFindClosingBracket(text, open, out var close);

        await Assert.That(ok).IsFalse();
        await Assert.That(close).IsEqualTo(-1);
    }
}
