using System.Text;
using System.Text.RegularExpressions;

namespace Danfma.MySheet.Excel;

/// <summary>
/// Rewrites a shared-formula master's text for a slave cell by shifting every RELATIVE cell reference by
/// the (row, column) delta between the two cells — the same expansion Excel performs for
/// <c>&lt;f t="shared"&gt;</c> groups. The shift happens at the TEXT level (not on the parsed tree)
/// because only the text still carries the <c>$</c> markers: absolute components must not move, and the
/// MySheet AST intentionally drops <c>$</c>. String literals (<c>"…"</c>) and quoted sheet names
/// (<c>'…'</c>) are copied untouched.
/// </summary>
internal static partial class SharedFormulaShifter
{
    [GeneratedRegex(@"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")]
    private static partial Regex CellRefPattern { get; }

    public static string Shift(string formula, string fromId, string toId)
    {
        var from = CellId.Parse(fromId);
        var to = CellId.Parse(toId);
        var deltaRow = to.Row - from.Row;
        var deltaColumn = to.Column - from.Column;

        if (deltaRow == 0 && deltaColumn == 0)
        {
            return formula;
        }

        var builder = new StringBuilder(formula.Length + 8);
        var index = 0;

        while (index < formula.Length)
        {
            var c = formula[index];

            if (c == '"' || c == '\'')
            {
                index = CopyQuoted(formula, index, builder);
            }
            else if (IsIdentifierChar(c) && (index == 0 || !IsIdentifierChar(formula[index - 1])))
            {
                index = CopyOrShiftToken(formula, index, builder, deltaRow, deltaColumn);
            }
            else
            {
                builder.Append(c);
                index++;
            }
        }

        return builder.ToString();
    }

    // Copies a quoted section verbatim, honoring the doubled-quote escape ("" or '').
    private static int CopyQuoted(string formula, int start, StringBuilder builder)
    {
        var quote = formula[start];
        var index = start + 1;

        builder.Append(quote);

        while (index < formula.Length)
        {
            builder.Append(formula[index]);

            if (formula[index] == quote)
            {
                if (index + 1 < formula.Length && formula[index + 1] == quote)
                {
                    builder.Append(quote);
                    index += 2;
                    continue;
                }

                return index + 1;
            }

            index++;
        }

        return index;
    }

    // Reads one identifier-like token and either shifts it (a standalone cell reference) or copies it as
    // is (function name, LET name, sheet name before '!', etc.).
    private static int CopyOrShiftToken(
        string formula,
        int start,
        StringBuilder builder,
        int deltaRow,
        int deltaColumn
    )
    {
        var end = start;

        while (end < formula.Length && IsIdentifierChar(formula[end]))
        {
            end++;
        }

        var token = formula[start..end];
        var next = end < formula.Length ? formula[end] : '\0';

        // A token followed by '(' is a function call and one followed by '!' is a sheet name — never a
        // cell reference, even when it happens to look like one (e.g. LOG10, Sheet2).
        var match = next is '(' or '!' ? null : CellRefPattern.Match(token);

        if (match is not { Success: true })
        {
            builder.Append(token);

            return end;
        }

        var columnAbsolute = match.Groups[1].Length > 0;
        var rowAbsolute = match.Groups[3].Length > 0;
        var column = ColumnNumber(match.Groups[2].Value);
        var row = int.Parse(match.Groups[4].Value);

        if (columnAbsolute)
        {
            builder.Append('$').Append(match.Groups[2].Value);
        }
        else
        {
            builder.Append(ColumnLetters(column + deltaColumn));
        }

        if (rowAbsolute)
        {
            builder.Append('$').Append(match.Groups[4].Value);
        }
        else
        {
            builder.Append(row + deltaRow);
        }

        return end;
    }

    private static bool IsIdentifierChar(char c) =>
        char.IsLetterOrDigit(c) || c is '_' or '.' or '$';

    private static int ColumnNumber(string letters)
    {
        var column = 0;

        foreach (var c in letters)
        {
            column = column * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }

        return column;
    }

    private static string ColumnLetters(int column)
    {
        var letters = string.Empty;

        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }

        return letters;
    }
}
