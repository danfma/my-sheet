using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

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

    private Criteria(Op op, bool numeric, double number, string text)
    {
        _op = op;
        _numeric = numeric;
        _number = number;
        _text = text;
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

        return _op switch
        {
            Op.Equal => WildcardMatch(_text, text),
            Op.NotEqual => !WildcardMatch(_text, text),
            _ => false,
        };
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
        var regex = new StringBuilder("^");

        foreach (var c in pattern)
        {
            regex.Append(
                c switch
                {
                    '*' => ".*",
                    '?' => ".",
                    _ => Regex.Escape(c.ToString()),
                }
            );
        }

        regex.Append('$');

        return Regex.IsMatch(text, regex.ToString(), RegexOptions.IgnoreCase);
    }
}
