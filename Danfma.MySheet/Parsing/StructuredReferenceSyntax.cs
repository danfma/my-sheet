using System.Text;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Parsing;

/// <summary>
/// The structured (table) reference syntax — the <c>[...]</c> suffix of <c>Tabela1[Valor]</c>,
/// <c>Tabela1[[#Headers],[Valor]]</c> and friends. Both halves of the syntax live here, the grammar
/// (decode) and its inverse (canonical rendering), so encode and decode cannot drift — the same reason
/// <c>FormulaWriter.IsSimpleSheetName</c> sits next to <c>WriteSheetQualifier</c>: the scanner, the
/// grammar (<c>Parse</c>) and the canonical renderer (<c>Write</c>) are one unit.
/// <para>
/// EVERY "measured" note in this file has the same provenance: <b>Aspose.Cells 26.6.0, PLAIN entry (never
/// <c>SetArrayFormula</c>), 2026-09-11</b>, over a <c>Data!Tabela1</c> at A1:P4 — one header row, three
/// equal data rows and an optional totals row — whose sixteen headers ARE the awkward names the notes quote
/// (<c>Valor</c>, <c>Sales Amount</c>, <c>#OfItems</c>, <c>" Padded "</c>, <c>e#f</c>, <c>Total (USD)</c>,
/// <c>% Comissao</c>, <c>Rev#1</c>, <c>Unit_Price</c>, <c>a,b</c>, <c>a:b</c>, a line-break name,
/// <c>a'b</c>, <c>c@d</c>, <c>x[y</c>, <c>p]q</c>), each column's three cells equal so a <c>SUM</c>
/// identifies the column. A note says "stored back" for what the oracle writes when the formula is read
/// again, which is how the canonical spelling was derived. The mode is named because it matters elsewhere
/// in this engine, not here: twenty rows re-measured ARRAY-ENTERED canonicalise and evaluate identically,
/// row for row, so the spelling rule is entry-mode independent.
/// </para>
/// </summary>
internal static class StructuredReferenceSyntax
{
    /// <summary>
    /// Decodes the <c>[...]</c> suffix of a structured reference into a <see cref="TableReference"/>.
    /// <paramref name="token"/> is the raw <see cref="TokenType.BracketedSpecifier"/> (brackets included,
    /// undecoded) and <paramref name="tableName"/> the identifier before it.
    /// <para>
    /// The payload is taken VERBATIM — there is no <c>Trim</c> on it, because Excel does not trim one:
    /// measured on Aspose.Cells 26.6.0 (PLAIN entry, 2026-09-11), <c>SUM(Tabela1[ Valor ])</c> over a header
    /// <c>"Valor"</c> is REJECTED (<c>Invalid table column:  Valor</c>) while the same spelling over a header
    /// <c>" Padded "</c> resolves, so the spaces belong to the name. Each ITEM of the item-list form is
    /// trimmed OUTSIDE its brackets, which is a different rule and is also measured
    /// (<c>Tabela1[[#Headers], [#Data], [Valor]]</c> resolves and is stored back without the spaces).
    /// </para>
    /// <para>
    /// Classification happens on the RAW text, BEFORE decoding: an item is a specifier iff its undecoded
    /// text starts with <c>#</c>. That ordering is what makes <c>Tabela1['#OfItems]</c> the column
    /// <c>#OfItems</c> — Microsoft's own example — instead of an unknown specifier.
    /// </para>
    /// </summary>
    internal static Expression Parse(string tableName, Token token)
    {
        var payload = token.Text[1..^1];
        var payloadStart = token.Position + 1;

        if (payload.Length == 0)
        {
            // Tabela1[] is ACCEPTED and means the whole DATA body: measured, Aspose reads it back (the
            // cell.Formula surface) as the bare `Tabela1` and, over a table WITH a totals row, answers the
            // data (90) where [#All] answers 180. A real xlsx carries the spelling too — the SAVED <f>
            // keeps `SUM(Tabela1[])`, so the two surfaces disagree and only this one names them.
            return new TableReference(tableName, null, TableArea.Data);
        }

        switch (payload[0])
        {
            case '@':
                // The current-row forms are valid Excel that MySheet does not model (S1 scope decision):
                // Tabela1[@Valor] outside the table evaluates to #VALUE!, it is not rejected.
                throw Unsupported(
                    $"The current-row form '{token.Text}' is not supported",
                    payloadStart,
                    token.Text
                );

            case '#':
                return new TableReference(
                    tableName,
                    null,
                    ParseSpecifier(payload, payloadStart, token.Text)
                );

            case '[':
                return ParseItemList(tableName, payload, payloadStart, token.Text);

            default:
                return new TableReference(
                    tableName,
                    DecodeName(payload, payloadStart, token.Text),
                    TableArea.Data
                );
        }
    }

    /// <summary>
    /// Renders a <see cref="TableReference"/> in Excel's CANONICAL spelling — the one the oracle stores back
    /// unchanged. Measured on Aspose.Cells 26.6.0, PLAIN entry, 2026-09-11, over a 16-column fixture:
    /// SINGLE bracket always, a <c>'</c> prefixed to each of <c>[ ] # ' @</c>, and the double-bracket item
    /// form for exactly one column shape — a name carrying leading or trailing whitespace. Every
    /// double-bracket spelling of an unpadded name is rewritten to the single-bracket one
    /// (<c>[[Valor]]</c> → <c>[Valor]</c>, <c>[['#OfItems]]</c> → <c>['#OfItems]</c>) and <c>[ Padded ]</c>
    /// is rewritten to <c>[[ Padded ]]</c>.
    /// <para>
    /// The table name goes in BARE, with no quoting branch: <c>Table.ValidateName</c> admits only letters,
    /// digits, <c>_</c> and <c>.</c> with a letter or <c>_</c> first, so a registered table name never needs
    /// quotes — and the quoted spelling is not Excel anyway (<c>SUM('Tabela1'[Valor])</c> is REJECTED,
    /// <c>Invalid "'"</c>). The ambient deltaRow/deltaColumn of a shared formula is ignored on purpose: a
    /// structured reference has no position component, which is also why it stays on the shared-master path.
    /// </para>
    /// </summary>
    internal static void Write(StringBuilder builder, TableReference reference)
    {
        var column = reference.ColumnName;
        var lastColumn = reference.LastColumnName;

        // #Data is implicit ONLY alongside a column, so T[Valor] rather than T[[#Data],[Valor]]; with no
        // column it must be written, because T[] is stored by Excel as the bare table name and MySheet has
        // no node for that spelling yet.
        string[] specifiers = reference.Area switch
        {
            TableArea.Data when column is not null => [],
            TableArea.Data => ["#Data"],
            TableArea.All => ["#All"],
            TableArea.Headers => ["#Headers"],
            TableArea.Totals => ["#Totals"],
            TableArea.HeadersAndData => ["#Headers", "#Data"],
            TableArea.DataAndTotals => ["#Data", "#Totals"],
            _ => throw new NotSupportedException(
                $"No Excel rendering for table area '{reference.Area}'."
            ),
        };

        // The bare-payload form carries exactly one part and no fencing: a lone specifier (T[#All]) or an
        // unpadded column name (T[Valor]). Everything else — a specifier pair, a specifier beside a column,
        // or a column whose edge whitespace the bare form would lose — brackets each part separately.
        var barePayload = column is null
            ? specifiers.Length == 1
            : specifiers.Length == 0 && lastColumn is null && !HasEdgeWhitespace(column);

        builder.Append(reference.TableName).Append('[');

        if (barePayload)
        {
            if (column is null)
            {
                builder.Append(specifiers[0]);
            }
            else
            {
                EscapeName(builder, column);
            }
        }
        else
        {
            var separator = "";

            foreach (var specifier in specifiers)
            {
                builder.Append(separator).Append('[').Append(specifier).Append(']');
                separator = ",";
            }

            if (column is not null)
            {
                builder.Append(separator).Append('[');
                EscapeName(builder, column);
                builder.Append(']');

                if (lastColumn is not null)
                {
                    builder.Append(":[");
                    EscapeName(builder, lastColumn);
                    builder.Append(']');
                }
            }
        }

        builder.Append(']');
    }

    private static Expression ParseItemList(
        string tableName,
        string payload,
        int payloadStart,
        string tokenText
    )
    {
        if (payload is "[]")
        {
            // Tabela1[[]] is ACCEPTED and means #All, not the data body: measured on two fixtures (180 over
            // a totals-row table where [#Data] answers 90) and stored back as [#All]. An empty item
            // ANYWHERE else is rejected below, because nothing measures it.
            return new TableReference(tableName, null, TableArea.All);
        }

        var specifiers = (First: TableArea.Data, Second: TableArea.Data, Count: 0);
        string? column = null;
        var (itemsPayload, lastColumn) = ExtractColumnSpan(payload, payloadStart, tokenText);

        foreach (var (item, offset) in SplitItems(itemsPayload))
        {
            var position = payloadStart + offset;

            if (!TryFindClosingBracket(item, 0, out var close) || close != item.Length - 1)
            {
                throw Invalid(
                    $"'{item}' is not a bracketed item of a structured reference",
                    position,
                    tokenText
                );
            }

            var inner = item[1..^1];
            var innerPosition = position + 1;

            if (inner.Length == 0)
            {
                throw Invalid("A structured reference item cannot be empty", position, tokenText);
            }

            switch (inner[0])
            {
                case '#':
                    var area = ParseSpecifier(inner, innerPosition, tokenText);

                    if (specifiers.Count == 2)
                    {
                        throw Invalid(
                            "A structured reference names at most two row specifiers",
                            innerPosition,
                            tokenText
                        );
                    }

                    specifiers =
                        specifiers.Count == 0
                            ? (area, TableArea.Data, 1)
                            : (specifiers.First, area, 2);
                    break;

                case '@':
                    throw Unsupported(
                        $"The current-row form '{tokenText}' is not supported",
                        innerPosition,
                        tokenText
                    );

                default:
                    if (column is not null)
                    {
                        throw Invalid(
                            "A structured reference names at most one column",
                            innerPosition,
                            tokenText
                        );
                    }

                    column = DecodeName(inner, innerPosition, tokenText);
                    break;
            }
        }

        if (lastColumn is not null && column is null)
        {
            throw Invalid("A column span needs a first column", payloadStart, tokenText);
        }

        return new TableReference(
            tableName,
            column,
            Combine(specifiers, payloadStart, tokenText),
            lastColumn
        );
    }

    private static (string ItemsPayload, string? LastColumn) ExtractColumnSpan(
        string payload,
        int payloadStart,
        string tokenText
    )
    {
        var depth = 0;

        for (var i = 0; i < payload.Length; i++)
        {
            switch (payload[i])
            {
                case '\'':
                    i++;
                    break;
                case '[':
                    depth++;
                    break;
                case ']':
                    depth--;
                    break;
                case ':' when depth == 0:
                    var right = payload[(i + 1)..].Trim();
                    if (
                        !TryFindClosingBracket(right, 0, out var close)
                        || close != right.Length - 1
                    )
                    {
                        throw Invalid(
                            "A column span needs a bracketed last column",
                            payloadStart + i + 1,
                            tokenText
                        );
                    }

                    if (right.Length > 2 && right[1] is '#' or '@')
                    {
                        throw Invalid(
                            "A column span endpoint must be a column name",
                            payloadStart + i + 2,
                            tokenText
                        );
                    }

                    var leftEnd = i;
                    while (leftEnd > 0 && char.IsWhiteSpace(payload[leftEnd - 1]))
                    {
                        leftEnd--;
                    }

                    return (
                        payload[..leftEnd],
                        DecodeName(right[1..^1], payloadStart + i + 2, tokenText)
                    );
            }
        }

        return (payload, null);
    }

    // Items are collected ORDER-INDEPENDENTLY because that is what the oracle does rather than reject:
    // measured, [[Valor],[#Data]] and [[#Data],[#Headers]] are both accepted and stored back reordered to
    // [[#Data],[Valor]] and [[#Headers],[#Data]]. The legal specifier SETS are only the four singletons
    // plus {Headers,Data} and {Data,Totals} — every other pair and every triple is rejected by the oracle
    // with "Specified rows to make up the contiguous range can only be one of following items: Headers,
    // Data, Totals, Data and Headers, Data and Totals, CurrentRow".
    private static TableArea Combine(
        (TableArea First, TableArea Second, int Count) specifiers,
        int payloadStart,
        string tokenText
    ) =>
        specifiers switch
        {
            { Count: 0 } => TableArea.Data,
            { Count: 1 } => specifiers.First,
            (TableArea.Headers, TableArea.Data, _) or (TableArea.Data, TableArea.Headers, _) =>
                TableArea.HeadersAndData,
            (TableArea.Data, TableArea.Totals, _) or (TableArea.Totals, TableArea.Data, _) =>
                TableArea.DataAndTotals,
            _ => throw Invalid(
                $"'{specifiers.First}' and '{specifiers.Second}' are not a contiguous pair of table rows",
                payloadStart,
                tokenText
            ),
        };

    // Splits the item-list payload at its TOP-LEVEL commas, with each item trimmed OUTSIDE its brackets and
    // its offset within the payload kept so a rejected item reports a position inside the brackets. A
    // nested bracket group is opaque here — where it ends is decided by the ONE scanner above, which is
    // what guarantees the splitter cannot disagree with the tokenizer about a payload's shape.
    private static List<(string Item, int Offset)> SplitItems(string payload)
    {
        var items = new List<(string, int)>(3);
        var start = 0;
        var i = 0;

        while (i < payload.Length)
        {
            switch (payload[i])
            {
                case '\'':
                    i += 2;
                    break;

                case '[':
                    i = TryFindClosingBracket(payload, i, out var close)
                        ? close + 1
                        : payload.Length;
                    break;

                case ',':
                    items.Add(Slice(payload, start, i));
                    start = ++i;
                    break;

                default:
                    i++;
                    break;
            }
        }

        items.Add(Slice(payload, start, payload.Length));

        return items;

        static (string Item, int Offset) Slice(string payload, int start, int end)
        {
            while (start < end && char.IsWhiteSpace(payload[start]))
            {
                start++;
            }

            while (end > start && char.IsWhiteSpace(payload[end - 1]))
            {
                end--;
            }

            return (payload[start..end], start);
        }
    }

    // Microsoft's documented specifier set, and the oracle's: #All / #Data / #Headers / #Totals, case
    // insensitive and canonicalised by the writer ([#totals] is stored back as [#Totals]).
    private static TableArea ParseSpecifier(string text, int position, string tokenText)
    {
        if (Is(text, "#All"))
        {
            return TableArea.All;
        }

        if (Is(text, "#Data"))
        {
            return TableArea.Data;
        }

        if (Is(text, "#Headers"))
        {
            return TableArea.Headers;
        }

        if (Is(text, "#Totals"))
        {
            return TableArea.Totals;
        }

        if (Is(text, "#This Row"))
        {
            throw Unsupported(
                $"The current-row form '{tokenText}' is not supported",
                position,
                tokenText
            );
        }

        throw Invalid($"Unknown table row specifier '{text}'", position, tokenText);

        static bool Is(string text, string specifier) =>
            string.Equals(text, specifier, StringComparison.OrdinalIgnoreCase);
    }

    // Excel's escape, measured (Aspose.Cells 26.6.0, PLAIN entry, 2026-09-11, both review probes): a `'`
    // makes the NEXT character literal ONLY when that character is one of the five specials — '' -> ',
    // '] -> ], '[ -> [, '# -> # and '@ -> @. Before an ordinary character the apostrophe is LITERAL:
    // `Tabela1[a'b]` resolves to the column named `a'b`, which is why the canonical writer doubles every
    // apostrophe. Runs AFTER classification (see Parse's remarks), never before, or a column named #OfItems
    // would decode into an unknown specifier.
    private static string DecodeName(string raw, int position, string tokenText)
    {
        if (!raw.Contains('\''))
        {
            return raw;
        }

        var builder = new StringBuilder(raw.Length);

        for (var i = 0; i < raw.Length; i++)
        {
            if (raw[i] != '\'')
            {
                builder.Append(raw[i]);
                continue;
            }

            if (i + 1 == raw.Length)
            {
                throw Invalid(
                    $"A dangling ' escape ends the table column name '{raw}'",
                    position + i,
                    tokenText
                );
            }

            var next = raw[i + 1];

            if (next is '[' or ']' or '#' or '\'' or '@')
            {
                builder.Append(next);
                i++;
            }
            else
            {
                builder.Append('\'');
            }
        }

        return builder.ToString();
    }

    // The ONE case whose canonical spelling is the double-bracket item form: the bare payload is taken
    // verbatim, so it CAN carry edge whitespace, but the oracle still fences such a name in its own
    // brackets ([ Padded ] is stored back as [[ Padded ]]). Inside an item list the column takes one
    // bracket pair even when padded, because there the outer bracket is already the list's.
    private static bool HasEdgeWhitespace(string name) =>
        name.Length > 0 && (char.IsWhiteSpace(name[0]) || char.IsWhiteSpace(name[^1]));

    // The escape set is Microsoft's documented one and matches what the oracle emits on the cell.Formula
    // READBACK surface: [ ] # ' @ and nothing else, so a comma, a colon, a parenthesis, a '%' and a raw
    // newline all pass through untouched ([e#f] -> [e'#f], [a'b] -> [a''b], [c@d] -> [c'@d], [a,b] and
    // [Line\nBreak] stored identically). The SAVED <f> disagrees for '@' — the oracle writes [c@d] and
    // [a@b] unescaped there while still escaping '#' — but this writer's output is accepted and
    // readback-stable on both surfaces, which is the invariant that matters here (measured,
    // Aspose.Cells 26.6.0, PLAIN, 2026-09-11, flash-review probe).
    private static void EscapeName(StringBuilder builder, string name)
    {
        foreach (var c in name)
        {
            if (c is '[' or ']' or '#' or '\'' or '@')
            {
                builder.Append('\'');
            }

            builder.Append(c);
        }
    }

    private static ParseException Invalid(string message, int position, string tokenText) =>
        new(ParseErrorKind.InvalidStructuredReference, message, position, tokenText);

    private static ParseException Unsupported(string message, int position, string tokenText) =>
        new(ParseErrorKind.UnsupportedStructuredReference, message, position, tokenText);

    /// <summary>
    /// Finds the <c>]</c> that balances the <c>[</c> at <paramref name="openIndex"/>, honouring Excel's
    /// <c>'</c> escape: an apostrophe and the character after it are skipped as a pair, so <c>']</c>,
    /// <c>'[</c>, <c>''</c>, <c>'#</c> and <c>'@</c> never count. Nothing else is interpreted — no decoding,
    /// no trimming, and no whitespace rule, so a raw newline inside the brackets is payload (Excel stores a
    /// column named with a line break as a real <c>\n</c> in the cell's formula). The ONE scanner shared by
    /// the tokenizer's reader, the item splitter and the per-item bracket check, which is what guarantees the
    /// tokenizer can never accept a payload the splitter mis-splits.
    /// </summary>
    /// <returns>
    /// <c>true</c> with <paramref name="closeIndex"/> at the balancing <c>]</c>; <c>false</c> (index -1) when
    /// the brackets never balance before the end of the text — including a trailing <c>'</c> that escapes
    /// past the end — or when <paramref name="openIndex"/> is out of range or not at a <c>[</c> (the item
    /// splitter relies on that answer to mean "not a bracketed item").
    /// </returns>
    internal static bool TryFindClosingBracket(string text, int openIndex, out int closeIndex)
    {
        closeIndex = -1;

        if ((uint)openIndex >= (uint)text.Length || text[openIndex] != '[')
        {
            return false;
        }

        var depth = 0;

        for (var i = openIndex; i < text.Length; i++)
        {
            switch (text[i])
            {
                case '\'':
                    i++; // the escapee, whatever it is; past the end it simply ends the scan unbalanced
                    break;
                case '[':
                    depth++;
                    break;
                case ']':
                    if (--depth == 0)
                    {
                        closeIndex = i;
                        return true;
                    }

                    break;
            }
        }

        return false;
    }
}
