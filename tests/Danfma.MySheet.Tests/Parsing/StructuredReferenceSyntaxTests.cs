using System.Text;
using Danfma.MySheet.Expressions;
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

    // === Item 6: the grammar (Parse) =====================================================================

    // Every accepted row below was measured on Aspose.Cells 26.6.0, PLAIN entry, over a 16-column `Tabela1`
    // whose headers are exactly the awkward names used here (`Valor`, `Sales Amount`, `#OfItems`,
    // `" Padded "`, `e#f`, `Total (USD)`, `% Comissao`, `a,b`, `a:b`, `Line\nBreak`, `a'b`, `c@d`, `x[y`,
    // `p]q`), each column's three data cells equal so the SUM identifies the column. Re-measured
    // 2026-09-11 for this task; the numbers live in the task report, the SHAPES live here.
    private static TableReference ParseSuffix(string suffix, int position = 7) =>
        (TableReference)
            StructuredReferenceSyntax.Parse(
                "Tabela1",
                new Token(TokenType.BracketedSpecifier, suffix, position)
            );

    // The payload is taken VERBATIM — no Trim. Aspose rejects `Tabela1[ Valor ]` over a header "Valor"
    // (`Invalid table column:  Valor`) and resolves it over a header `" Padded "`, so the spaces are part
    // of the name; a trim would turn a miss into a silent hit on a different column.
    // B1's ordering rule: an item is a SPECIFIER iff its RAW (undecoded) text starts with `#`, so
    // `['#OfItems]` is the COLUMN `#OfItems` and not an unknown specifier.
    [Test]
    [Arguments("[Valor]", "Valor", TableArea.Data)]
    [Arguments("[[Valor]]", "Valor", TableArea.Data)]
    [Arguments("[Sales Amount]", "Sales Amount", TableArea.Data)]
    [Arguments("[[Sales Amount]]", "Sales Amount", TableArea.Data)]
    [Arguments("[Total (USD)]", "Total (USD)", TableArea.Data)]
    [Arguments("[[Total (USD)]]", "Total (USD)", TableArea.Data)]
    [Arguments("['#OfItems]", "#OfItems", TableArea.Data)]
    [Arguments("[['#OfItems]]", "#OfItems", TableArea.Data)]
    [Arguments("[e'#f]", "e#f", TableArea.Data)]
    [Arguments("[e#f]", "e#f", TableArea.Data)] // unescaped `#` after the first char is still a name char
    [Arguments("[a''b]", "a'b", TableArea.Data)]
    [Arguments("[c'@d]", "c@d", TableArea.Data)]
    [Arguments("[x'[y]", "x[y", TableArea.Data)]
    [Arguments("[p']q]", "p]q", TableArea.Data)]
    [Arguments("[ Padded ]", " Padded ", TableArea.Data)]
    [Arguments("[[ Padded ]]", " Padded ", TableArea.Data)]
    [Arguments("[a,b]", "a,b", TableArea.Data)] // a comma is only a separator in the item-list form
    [Arguments("[a:b]", "a:b", TableArea.Data)] // ...and so is a colon
    [Arguments("[Line\nBreak]", "Line\nBreak", TableArea.Data)]
    public async Task Parse_ColumnOnly(string suffix, string column, TableArea area)
    {
        var reference = ParseSuffix(suffix);

        await Assert.That(reference.TableName).IsEqualTo("Tabela1");
        await Assert.That(reference.ColumnName).IsEqualTo(column);
        await Assert.That(reference.Area).IsEqualTo(area);
    }

    // `Tabela1[]` is ACCEPTED and means the whole DATA body, measured: Aspose reads it back on the
    // cell.Formula surface as the bare `Tabela1` (the SAVED <f> keeps `SUM(Tabela1[])` — the surfaces
    // disagree) and, over a table WITH a totals row, answers 90 (the data) where `[#All]` answers 180.
    // `Tabela1[[]]` is also accepted but means `#All` — measured on the same two fixtures (180 with the
    // totals row) and stored back as `=SUM(Tabela1[#All])`.
    [Test]
    [Arguments("[]", TableArea.Data)]
    [Arguments("[[]]", TableArea.All)]
    [Arguments("[#All]", TableArea.All)]
    [Arguments("[#Data]", TableArea.Data)]
    [Arguments("[#Headers]", TableArea.Headers)]
    [Arguments("[#Totals]", TableArea.Totals)]
    [Arguments("[#totals]", TableArea.Totals)] // case-insensitive, canonicalised by the writer
    [Arguments("[#DATA]", TableArea.Data)]
    [Arguments("[[#Data]]", TableArea.Data)]
    [Arguments("[[#Headers],[#Data]]", TableArea.HeadersAndData)]
    [Arguments("[[#Data],[#Totals]]", TableArea.DataAndTotals)]
    [Arguments("[[#Data],[#Headers]]", TableArea.HeadersAndData)] // Aspose REORDERS rather than rejecting
    [Arguments("[[#Totals],[#Data]]", TableArea.DataAndTotals)]
    public async Task Parse_AreaWithNoColumn(string suffix, TableArea area)
    {
        var reference = ParseSuffix(suffix);

        await Assert.That(reference.ColumnName).IsNull();
        await Assert.That(reference.Area).IsEqualTo(area);
    }

    // Items are collected ORDER-INDEPENDENTLY, because that is what the oracle does: `[[Valor],[#Data]]`
    // and `[[Valor],[#Headers]]` are both ACCEPTED and stored back reordered to `[[#Data],[Valor]]` /
    // `[[#Headers],[Valor]]`. A space after the separator is tolerated and stripped.
    [Test]
    [Arguments("[[#All],[Valor]]", "Valor", TableArea.All)]
    [Arguments("[[#Data],[Valor]]", "Valor", TableArea.Data)]
    [Arguments("[[Valor],[#Data]]", "Valor", TableArea.Data)]
    [Arguments("[[Valor],[#Headers]]", "Valor", TableArea.Headers)]
    [Arguments("[[#Headers],[#Data],[Valor]]", "Valor", TableArea.HeadersAndData)]
    [Arguments("[[#Headers], [#Data], [Valor]]", "Valor", TableArea.HeadersAndData)]
    [Arguments("[[#Headers],[#Data],[% Comissao]]", "% Comissao", TableArea.HeadersAndData)]
    [Arguments("[[#Headers],[#Data],[ Padded ]]", " Padded ", TableArea.HeadersAndData)]
    [Arguments("[[#Data],[#Totals],[Valor]]", "Valor", TableArea.DataAndTotals)]
    [Arguments("[[#Data],[a,b]]", "a,b", TableArea.Data)] // the separator scan is bracket-depth aware
    [Arguments("[[#All],['#OfItems]]", "#OfItems", TableArea.All)]
    public async Task Parse_AreaWithColumn(string suffix, string column, TableArea area)
    {
        var reference = ParseSuffix(suffix);

        await Assert.That(reference.ColumnName).IsEqualTo(column);
        await Assert.That(reference.Area).IsEqualTo(area);
    }

    // Balanced brackets whose content the oracle REJECTS. Every row was measured: two columns is
    // `Unknown token with bracket` / `Invalid table column`, an unknown specifier is
    // `Invalid table rows type`, and the illegal specifier combinations answer "Specified rows to make up
    // the contiguous range can only be one of following items: Headers, Data, Totals, Data and Headers,
    // Data and Totals, CurrentRow". So the kind is Invalid, NOT Unsupported — the design's "classify
    // Unsupported when unsure whether Excel accepts a shape" tie-break no longer applies to any of them.
    [Test]
    [Arguments("[[A],[B]]")]
    [Arguments("[[Valor],[Sales Amount]]")]
    [Arguments("[#Bogus]")]
    [Arguments("[[#Bogus]]")]
    [Arguments("[[#All],[#Data]]")]
    [Arguments("[[#Headers],[#Totals]]")]
    [Arguments("[[#All],[#Headers]]")]
    [Arguments("[[#Headers],[#Data],[#Totals]]")]
    [Arguments("[[Valor]x]")] // an item whose brackets do not span it
    [Arguments("[[Valor]]]")] // ...nor this one
    [Arguments("[[],[Valor]]")] // an empty item anywhere but the whole payload
    [Arguments("[[#Data],]")]
    [Arguments("[[#Data],[Valor],[Outra]]")]
    public async Task Parse_Invalid(string suffix)
    {
        var exception = Assert.Throws<ParseException>(() => ParseSuffix(suffix));

        await Assert.That(exception!.Kind).IsEqualTo(ParseErrorKind.InvalidStructuredReference);
    }

    // Valid Excel that MySheet does not model — a scope decision, not parity. The current-row forms are
    // ACCEPTED by the oracle (`Tabela1[@Valor]` evaluates to #VALUE! outside the table and the saved xlsx
    // carries `[[#This Row],[Valor]]` for both spellings).
    [Test]
    [Arguments("[@Valor]")]
    [Arguments("[@]")]
    [Arguments("[#This Row]")]
    [Arguments("[#this row]")]
    [Arguments("[[#This Row],[Valor]]")]
    [Arguments("[[@],[Valor]]")]
    public async Task Parse_Unsupported(string suffix)
    {
        var exception = Assert.Throws<ParseException>(() => ParseSuffix(suffix));

        await Assert.That(exception!.Kind).IsEqualTo(ParseErrorKind.UnsupportedStructuredReference);
    }

    [Test]
    [Arguments("[[Valor]:[#Data]]")]
    public async Task Parse_InvalidSpanEndpoint(string suffix)
    {
        var exception = Assert.Throws<ParseException>(() => ParseSuffix(suffix));

        await Assert.That(exception!.Kind).IsEqualTo(ParseErrorKind.InvalidStructuredReference);
    }

    // A trailing `'` with nothing to escape. UNREACHABLE from the tokenizer — the scanner skips the escape
    // and its escapee as a pair, so such a payload eats its own `]` and the token never balances
    // (`Tabela1[Valor']` is UnterminatedBracketedReference, and the oracle rejects it too, as
    // `Invalid '['`). The guard exists for a hand-built token, which is exactly what this test supplies.
    [Test]
    [Arguments("[Valor']")]
    public async Task Parse_DanglingEscape_IsInvalid(string suffix)
    {
        var exception = Assert.Throws<ParseException>(() => ParseSuffix(suffix));

        await Assert.That(exception!.Kind).IsEqualTo(ParseErrorKind.InvalidStructuredReference);
        await Assert.That(exception.Message).Contains("escape");
    }

    // Every throw reports `token.Position + offset`, so the position lands INSIDE the brackets, on the
    // offending item rather than on the `[`. The token here starts at 7, as in `Tabela1[...]`.
    [Test]
    [Arguments("[#Bogus]", 8)]
    [Arguments("[@Valor]", 8)]
    [Arguments("[[#Bogus]]", 9)]
    [Arguments("[[#Data],[#Bogus]]", 17)] // the SECOND item's specifier, past the separator
    [Arguments("[[#Headers], [#Bogus]]", 21)] // ...and the decorative space does not shift it
    [Arguments("[[Valor],[Outra]]", 17)] // the second COLUMN is what is rejected, so its own offset
    public async Task Parse_ReportsAPositionInsideTheBrackets(string suffix, int position)
    {
        var exception = Assert.Throws<ParseException>(() => ParseSuffix(suffix));

        await Assert.That(exception!.Position).IsEqualTo(position);
    }

    // === Item 12: the writer (Write) =====================================================================

    // Measured canonical spelling, Aspose.Cells 26.6.0, PLAIN entry, re-measured 2026-09-11 on a fresh
    // 16-column fixture: SINGLE bracket always, a `'` prefixed to each of `[ ] # ' @`, and DOUBLE brackets
    // for exactly one case — a column name carrying leading or trailing whitespace. Every double-bracket
    // spelling of an unpadded name is rewritten to the single-bracket one (`[[Valor]]` → `[Valor]`,
    // `[['#OfItems]]` → `['#OfItems]`, `[[Total (USD)]]` → `[Total (USD)]`), while `[ Padded ]` is
    // rewritten to `[[ Padded ]]`. Inside an item list the column takes ONE bracket even when padded
    // (`[[#Headers],[#Data],[ Padded ]]`, stored identically), because there the outer bracket is the list.
    [Test]
    [Arguments(null, TableArea.Data, "Tabela1[#Data]")]
    [Arguments(null, TableArea.All, "Tabela1[#All]")]
    [Arguments(null, TableArea.Headers, "Tabela1[#Headers]")]
    [Arguments(null, TableArea.Totals, "Tabela1[#Totals]")]
    [Arguments(null, TableArea.HeadersAndData, "Tabela1[[#Headers],[#Data]]")]
    [Arguments(null, TableArea.DataAndTotals, "Tabela1[[#Data],[#Totals]]")]
    [Arguments("Valor", TableArea.Data, "Tabela1[Valor]")]
    [Arguments("Sales Amount", TableArea.Data, "Tabela1[Sales Amount]")]
    [Arguments("Total (USD)", TableArea.Data, "Tabela1[Total (USD)]")]
    [Arguments("% Comissao", TableArea.Data, "Tabela1[% Comissao]")]
    [Arguments("Unit_Price", TableArea.Data, "Tabela1[Unit_Price]")]
    [Arguments("#OfItems", TableArea.Data, "Tabela1['#OfItems]")]
    [Arguments("e#f", TableArea.Data, "Tabela1[e'#f]")]
    [Arguments("a'b", TableArea.Data, "Tabela1[a''b]")]
    [Arguments("c@d", TableArea.Data, "Tabela1[c'@d]")]
    [Arguments("x[y", TableArea.Data, "Tabela1[x'[y]")]
    [Arguments("p]q", TableArea.Data, "Tabela1[p']q]")]
    [Arguments("a,b", TableArea.Data, "Tabela1[a,b]")] // a comma is NOT in the escape set
    [Arguments("a:b", TableArea.Data, "Tabela1[a:b]")] // ...and neither is a colon
    [Arguments("Line\nBreak", TableArea.Data, "Tabela1[Line\nBreak]")]
    [Arguments(" Padded ", TableArea.Data, "Tabela1[[ Padded ]]")] // the ONE double-bracket case
    [Arguments("Trailing ", TableArea.Data, "Tabela1[[Trailing ]]")]
    [Arguments("Valor", TableArea.All, "Tabela1[[#All],[Valor]]")]
    [Arguments("Valor", TableArea.Headers, "Tabela1[[#Headers],[Valor]]")]
    [Arguments("Valor", TableArea.Totals, "Tabela1[[#Totals],[Valor]]")]
    [Arguments("Valor", TableArea.HeadersAndData, "Tabela1[[#Headers],[#Data],[Valor]]")]
    [Arguments("Valor", TableArea.DataAndTotals, "Tabela1[[#Data],[#Totals],[Valor]]")]
    [Arguments(" Padded ", TableArea.HeadersAndData, "Tabela1[[#Headers],[#Data],[ Padded ]]")]
    [Arguments("#OfItems", TableArea.All, "Tabela1[[#All],['#OfItems]]")]
    public async Task Write_RendersTheCanonicalSpelling(
        string? column,
        TableArea area,
        string expected
    )
    {
        var builder = new StringBuilder();

        StructuredReferenceSyntax.Write(builder, new TableReference("Tabela1", column, area));

        await Assert.That(builder.ToString()).IsEqualTo(expected);
    }

    // The round-trip invariant the single-file design exists to protect: for every node the PARSER can
    // produce, re-parsing the canonical text yields the SAME node. Driven from the same column names as the
    // rows above so a spelling change cannot be made in one half only. The qualifier matters: a node built
    // by hand with an EMPTY column name is outside the invariant, because the writer has no spelling that
    // distinguishes it from "no column" (`T[]` is the whole data body) and the grammar never produces one.
    [Test]
    [MethodDataSource(nameof(EveryShape))]
    public async Task Write_ThenParse_IsTheSameNode(string? column, TableArea area)
    {
        var reference = new TableReference("Tabela1", column, area);
        var builder = new StringBuilder();
        StructuredReferenceSyntax.Write(builder, reference);
        var written = builder.ToString();
        var open = written.IndexOf('[');

        var reparsed = ParseSuffix(written[open..], open);

        await Assert.That(reparsed).IsEqualTo(reference);
    }

    // Exhaustive property pass over every payload the tokenizer can produce out of the ten characters the
    // grammar reacts to, up to four of them — the alphabet is too small to spell a specifier (#Data needs
    // five ordered characters the shuffler will not often produce), so coverage of the specifier forms
    // rests on the hand-picked rows above and this pass guards the PARTITION: no third exception kind, no
    // round-trip hole. Two invariants, both of which a hand-picked row list can miss:
    // (a) the only exception that escapes Parse is a ParseException carrying one of the three structured
    // kinds — never an IndexOutOfRange from the offsets, a NullReference, or a raw InvalidOperation; and
    // (b) every ACCEPTED payload's canonical rendering re-parses to the same node, so no spelling the
    // grammar admits has a round-trip hole. The filter keeps exactly the strings TryFindClosingBracket
    // accepts as one whole token, which is what the tokenizer hands Parse. 4869 of the generated
    // candidates survive that filter, 3598 accepted by the grammar and 1271 rejected — both asserted below.
    [Test]
    public async Task Parse_OverEveryWellFormedPayload_EitherThrowsAStructuredKind_OrRoundTrips()
    {
        const string alphabet = "[]#'@,: aD";
        var accepted = 0;
        var rejected = 0;
        var failures = new List<string>();

        foreach (var text in Candidates(alphabet, 4))
        {
            if (
                !StructuredReferenceSyntax.TryFindClosingBracket(text, 0, out var close)
                || close != text.Length - 1
            )
            {
                continue;
            }

            TableReference reference;

            try
            {
                reference = ParseSuffix(text);
                accepted++;
            }
            catch (ParseException exception)
            {
                rejected++;

                if (
                    exception.Kind
                    is not (
                        ParseErrorKind.InvalidStructuredReference
                        or ParseErrorKind.UnsupportedStructuredReference
                    )
                )
                {
                    failures.Add($"{text} -> wrong kind {exception.Kind}");
                }

                continue;
            }
            catch (Exception exception)
            {
                failures.Add($"{text} -> {exception.GetType().Name}: {exception.Message}");
                continue;
            }

            var builder = new StringBuilder();
            StructuredReferenceSyntax.Write(builder, reference);
            var written = builder.ToString();
            var open = written.IndexOf('[');

            try
            {
                var reparsed = ParseSuffix(written[open..], open);

                if (reparsed != reference)
                {
                    failures.Add($"{text} -> {written} -> {reparsed}, expected {reference}");
                }
            }
            catch (Exception exception)
            {
                failures.Add($"{text} -> {written} -> {exception.GetType().Name}");
            }
        }

        await Assert.That(failures).IsEmpty();

        // The exact split, so the pass can never go vacuous: a filter or an alphabet edit that stopped
        // generating payloads would otherwise leave an empty loop asserting nothing.
        await Assert.That(accepted).IsEqualTo(3598);
        await Assert.That(rejected).IsEqualTo(1271);
    }

    // "[", then every string of up to maxBody alphabet characters, then "]".
    private static IEnumerable<string> Candidates(string alphabet, int maxBody)
    {
        var bodies = new List<string> { "" };

        for (var length = 0; length < maxBody; length++)
        {
            var next = new List<string>(bodies.Count * alphabet.Length);

            foreach (var body in bodies.Where(body => body.Length == length))
            {
                foreach (var c in alphabet)
                {
                    next.Add(body + c);
                }
            }

            bodies.AddRange(next);
        }

        return bodies.Select(body => $"[{body}]");
    }

    public static IEnumerable<(string? Column, TableArea Area)> EveryShape()
    {
        string?[] columns =
        [
            null,
            "Valor",
            "Sales Amount",
            "Total (USD)",
            "#OfItems",
            "e#f",
            "a'b",
            "c@d",
            "x[y",
            "p]q",
            "a,b",
            "a:b",
            " Padded ",
            "Line\nBreak",
        ];

        foreach (var column in columns)
        {
            foreach (var area in Enum.GetValues<TableArea>())
            {
                yield return (column, area);
            }
        }
    }
}
