using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Danfma.MySheet.Expressions.Text;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// A COUNTIF/SUMIF-style criterion (e.g. <c>"&gt;5"</c>, <c>"&lt;&gt;x"</c>, <c>"apple*"</c>). Numeric
/// criteria compare numerically; text criteria match case-insensitively with <c>*</c>/<c>?</c> wildcards.
/// </summary>
internal sealed class Criteria
{
    private enum Op
    {
        Equal,
        NotEqual,
        Greater,
        Less,
        GreaterOrEqual,
        LessOrEqual,
    }

    private readonly Op _op;
    private readonly bool _numeric;
    private readonly double _number;
    private readonly string _text;

    // A text criterion's wildcard pattern resolved ONCE per criterion (null for numeric criteria), via the
    // shared RegexCache. The old code rebuilt the "^…$" pattern string and re-parsed the regex on EVERY cell —
    // for a *IFS scan over a 50k range that was a per-cell StringBuilder + pattern string, the dominant
    // transient once the range lists were gone. Routing through RegexCache also means the SAME text criterion
    // reused across cells/formulas (a common *IFS shape) shares one compiled Regex instead of recompiling it
    // once per Criteria.Parse.
    private readonly Regex? _regex;

    private Criteria(Op op, bool numeric, double number, string text)
    {
        _op = op;
        _numeric = numeric;
        _number = number;
        _text = text;
        _regex = numeric
            ? null
            : RegexCache.Get(BuildWildcardPattern(text), RegexOptions.IgnoreCase);
    }

    public static Criteria Parse(in ComputedValue value)
    {
        if (value.TryGetNumber(out var d))
        {
            return new Criteria(Op.Equal, numeric: true, d, string.Empty);
        }

        if (value.TryGetBoolean(out var b))
        {
            return new Criteria(Op.Equal, numeric: false, 0, b ? "TRUE" : "FALSE");
        }

        var (op, rest) = SplitOperator(value.TryGetText(out var s) ? s : string.Empty);

        return double.TryParse(rest, NumberStyles.Any, CultureInfo.InvariantCulture, out var number)
            ? new Criteria(op, numeric: true, number, rest)
            : new Criteria(op, numeric: false, 0, rest);
    }

    /// <summary>
    /// True when this is a plain numeric equality (<c>=k</c> / a bare number), the one criterion shape the
    /// Layer-2 numeric-equality map answers in O(1). A non-numeric equality, or any comparison/wildcard,
    /// returns false and the caller scans the cached snapshot linearly (still no cell re-read).
    /// </summary>
    internal bool TryGetNumericEquality(out double number)
    {
        if (_numeric && _op == Op.Equal)
        {
            number = _number;
            return true;
        }

        number = 0;
        return false;
    }

    public bool Matches(in ComputedValue cellValue)
    {
        if (_numeric)
        {
            if (!cellValue.TryGetNumber(out var cell))
            {
                // A non-numeric cell only satisfies a numeric criterion under "<>".
                return _op == Op.NotEqual;
            }

            return _op switch
            {
                Op.Equal => cell == _number,
                Op.NotEqual => cell != _number,
                Op.Greater => cell > _number,
                Op.Less => cell < _number,
                Op.GreaterOrEqual => cell >= _number,
                Op.LessOrEqual => cell <= _number,
                _ => false,
            };
        }

        var text = CellText(cellValue);

        try
        {
            return _op switch
            {
                Op.Equal => _regex!.IsMatch(text),
                Op.NotEqual => !_regex!.IsMatch(text),
                _ => false,
            };
        }
        catch (RegexMatchTimeoutException)
        {
            // The cached regex carries RegexCache's defensive 1s timeout (a wildcard criterion is not
            // normally at risk, but a formula can still chain enough '*' to backtrack pathologically).
            // Matches has no error channel, so a timeout fails safe as "does not match" — the same call
            // a caller would make for an unmatched cell.
            return false;
        }
    }

    private static string CellText(in ComputedValue value)
    {
        if (value.TryGetText(out var s))
        {
            return s;
        }

        if (value.TryGetNumber(out var d))
        {
            return d.ToString(CultureInfo.InvariantCulture);
        }

        if (value.TryGetBoolean(out var b))
        {
            return b ? "TRUE" : "FALSE";
        }

        return string.Empty;
    }

    private static (Op, string) SplitOperator(string s)
    {
        if (s.StartsWith(">=", StringComparison.Ordinal))
            return (Op.GreaterOrEqual, s[2..]);
        if (s.StartsWith("<=", StringComparison.Ordinal))
            return (Op.LessOrEqual, s[2..]);
        if (s.StartsWith("<>", StringComparison.Ordinal))
            return (Op.NotEqual, s[2..]);
        if (s.StartsWith(">", StringComparison.Ordinal))
            return (Op.Greater, s[1..]);
        if (s.StartsWith("<", StringComparison.Ordinal))
            return (Op.Less, s[1..]);
        if (s.StartsWith("=", StringComparison.Ordinal))
            return (Op.Equal, s[1..]);
        return (Op.Equal, s);
    }

    internal static bool WildcardMatch(string pattern, string text)
    {
        try
        {
            return BuildWildcardRegex(pattern).IsMatch(text);
        }
        catch (RegexMatchTimeoutException)
        {
            // Same fail-safe as Matches: no error channel here (XLOOKUP/XMATCH/LOOKUP wildcard scans treat
            // this as a per-element bool), so a timeout is reported as "no match" rather than propagating.
            return false;
        }
    }

    /// <summary>
    /// Resolves a wildcard pattern's compiled <see cref="Regex"/> ONCE, via the shared <see cref="RegexCache"/>
    /// (so the SAME pattern reused across cells/formulas shares one compiled instance, and a repeat lookup
    /// skips the string-building below entirely). The building block for a caller that scans many cells
    /// against the SAME pattern — e.g. <c>LookupMatching</c>'s wildcard mode — which should call this ONCE
    /// before the per-cell scan rather than resolving the pattern string on every <see cref="WildcardMatch"/>
    /// call (the RegexCache lookup avoids recompiling, but the "^…$" pattern string itself was still being
    /// rebuilt with a fresh StringBuilder per cell before this was hoisted out).
    /// </summary>
    internal static Regex BuildWildcardRegex(string pattern) =>
        RegexCache.Get(BuildWildcardPattern(pattern), RegexOptions.IgnoreCase);

    // Translates an Excel wildcard pattern (* → any run, ? → any single char, everything else literal) into
    // an anchored .NET regex pattern. Called ONCE per text criterion (see <see cref="_regex"/>), not per cell.
    private static string BuildWildcardPattern(string pattern)
    {
        var builder = new StringBuilder("^");

        foreach (var c in pattern)
        {
            builder.Append(
                c switch
                {
                    '*' => ".*",
                    '?' => ".",
                    _ => Regex.Escape(c.ToString()),
                }
            );
        }

        builder.Append('$');

        return builder.ToString();
    }
}
