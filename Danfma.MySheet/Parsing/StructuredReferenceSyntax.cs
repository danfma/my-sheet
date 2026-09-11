namespace Danfma.MySheet.Parsing;

/// <summary>
/// The structured (table) reference syntax — the <c>[...]</c> suffix of <c>Tabela1[Valor]</c>,
/// <c>Tabela1[[#Headers],[Valor]]</c> and friends. Both halves of the syntax live here, the grammar
/// (decode) and its inverse (canonical rendering), so encode and decode cannot drift — the same reason
/// <c>FormulaWriter.IsSimpleSheetName</c> sits next to <c>WriteSheetQualifier</c>. This file currently
/// carries the scanner; Phase 4 T2 adds <c>Parse</c> and <c>Write</c>.
/// </summary>
internal static class StructuredReferenceSyntax
{
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
