using System.Text;

namespace MySheet.Expressions;

/// <summary>
/// Translates the common Excel date/time format codes to .NET custom date-format strings, including the
/// month-vs-minute disambiguation of <c>m</c>/<c>mm</c>. Covers y, m, d, h, s, AM/PM and literals; the
/// full Excel format spec (sections, colours, fractions) is out of scope.
/// </summary>
internal static class ExcelDateFormat
{
    public static bool IsDateOrTime(string format)
    {
        var inQuote = false;

        foreach (var c in format)
        {
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && c is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M')
            {
                return true;
            }
        }

        return false;
    }

    public static string ToDotNet(string format)
    {
        var twelveHour = ContainsIgnoreCase(format, "AM/PM") || ContainsIgnoreCase(format, "A/P");
        var result = new StringBuilder();
        var i = 0;

        while (i < format.Length)
        {
            if (StartsWithIgnoreCase(format, i, "AM/PM"))
            {
                result.Append("tt");
                i += 5;
                continue;
            }

            if (StartsWithIgnoreCase(format, i, "A/P"))
            {
                result.Append("tt");
                i += 3;
                continue;
            }

            var c = format[i];

            if (c == '"')
            {
                i++;
                while (i < format.Length && format[i] != '"')
                {
                    AppendLiteral(result, format[i]);
                    i++;
                }

                i++; // closing quote
                continue;
            }

            if (c == '\\' && i + 1 < format.Length)
            {
                AppendLiteral(result, format[i + 1]);
                i += 2;
                continue;
            }

            if (char.IsLetter(c))
            {
                var lower = char.ToLowerInvariant(c);
                var run = 1;
                while (i + run < format.Length && char.ToLowerInvariant(format[i + run]) == lower)
                {
                    run++;
                }

                AppendToken(result, format, i, lower, run, twelveHour);
                i += run;
                continue;
            }

            // Separators (/ : - space) pass through; the invariant culture keeps them literal.
            result.Append(c);
            i++;
        }

        return result.ToString();
    }

    private static void AppendToken(StringBuilder result, string format, int index, char lower, int run, bool twelveHour)
    {
        switch (lower)
        {
            case 'y':
                result.Append(run >= 3 ? "yyyy" : "yy");
                break;

            case 'd':
                result.Append(run >= 4 ? "dddd" : run == 3 ? "ddd" : run >= 2 ? "dd" : "d");
                break;

            case 'h':
                var hour = twelveHour ? "h" : "H";
                result.Append(run >= 2 ? hour + hour : hour);
                break;

            case 's':
                result.Append(run >= 2 ? "ss" : "s");
                break;

            case 'm' when IsMinute(format, index, index + run):
                result.Append(run >= 2 ? "mm" : "m");
                break;

            case 'm':
                result.Append(run >= 4 ? "MMMM" : run == 3 ? "MMM" : run >= 2 ? "MM" : "M");
                break;

            default:
                for (var k = 0; k < run; k++)
                {
                    AppendLiteral(result, format[index + k]);
                }

                break;
        }
    }

    // 'm' means minutes when, ignoring separators, it follows an hour token or precedes a seconds token.
    private static bool IsMinute(string format, int start, int afterEnd)
    {
        var before = start - 1;
        while (before >= 0 && !char.IsLetter(format[before]))
        {
            before--;
        }

        if (before >= 0 && format[before] is 'h' or 'H')
        {
            return true;
        }

        var after = afterEnd;
        while (after < format.Length && !char.IsLetter(format[after]))
        {
            after++;
        }

        return after < format.Length && format[after] is 's' or 'S';
    }

    private static void AppendLiteral(StringBuilder result, char c)
    {
        if (char.IsLetter(c))
        {
            result.Append('\\');
        }

        result.Append(c);
    }

    private static bool ContainsIgnoreCase(string text, string value) =>
        text.Contains(value, StringComparison.OrdinalIgnoreCase);

    private static bool StartsWithIgnoreCase(string text, int index, string value) =>
        index + value.Length <= text.Length &&
        string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
}
