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

    public static Criteria Parse(object? value)
    {
        switch (value)
        {
            case double d:
                return new Criteria(Op.Equal, numeric: true, d, string.Empty);
            case bool b:
                return new Criteria(Op.Equal, numeric: false, 0, b ? "TRUE" : "FALSE");
            case null:
                return new Criteria(Op.Equal, numeric: false, 0, string.Empty);
        }

        var (op, rest) = SplitOperator(value.ToString() ?? string.Empty);

        return double.TryParse(rest, NumberStyles.Any, CultureInfo.InvariantCulture, out var number)
            ? new Criteria(op, numeric: true, number, rest)
            : new Criteria(op, numeric: false, 0, rest);
    }

    public bool Matches(object? cellValue)
    {
        if (_numeric)
        {
            if (cellValue is not double cell)
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
            _ => false, // ordering is not defined for text criteria
        };
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

    private static string CellText(object? value) =>
        value switch
        {
            null => string.Empty,
            string s => s,
            double d => d.ToString(CultureInfo.InvariantCulture),
            bool b => b ? "TRUE" : "FALSE",
            _ => string.Empty,
        };

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
