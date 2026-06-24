using System.Globalization;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Coerces the loosely-typed values returned by <see cref="Expression.Compute"/>
/// (raw CLR scalars or an <see cref="ErrorValue"/> node) into numbers, propagating errors.
/// </summary>
internal static class ValueCoercion
{
    /// <summary>
    /// Tries to coerce a computed value into a number. Returns <c>null</c> on success (with the
    /// number in <paramref name="number"/>), or the <see cref="ErrorValue"/> to propagate on failure.
    /// Blank/null coerces to 0; booleans to 1/0; numeric strings parse with the invariant culture.
    /// </summary>
    public static ErrorValue? TryToNumber(object? value, out double number)
    {
        switch (value)
        {
            case null:
                number = 0;
                return null;

            case double d:
                number = d;
                return null;

            case bool b:
                number = b ? 1 : 0;
                return null;

            case string s
                when double.TryParse(
                    s,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out var parsed
                ):
                number = parsed;
                return null;

            case ErrorValue error:
                number = 0;
                return error;

            default:
                number = 0;
                return ErrorValue.NotValue;
        }
    }

    /// <summary>
    /// Tries to coerce a computed value into a boolean condition (Excel truthiness). Returns <c>null</c>
    /// on success, or the <see cref="ErrorValue"/> to propagate. Blank→false; number→(≠0); text→#VALUE!.
    /// </summary>
    public static ErrorValue? TryToBool(object? value, out bool result)
    {
        switch (value)
        {
            case null:
                result = false;
                return null;

            case bool b:
                result = b;
                return null;

            case double d:
                result = d != 0;
                return null;

            case ErrorValue error:
                result = false;
                return error;

            default:
                result = false;
                return ErrorValue.NotValue;
        }
    }

    /// <summary>
    /// Coerces a computed value into its text representation. Returns <c>null</c> on success, or the
    /// <see cref="ErrorValue"/> to propagate. Blank→""; number→invariant string; bool→TRUE/FALSE.
    /// </summary>
    public static ErrorValue? TryToText(object? value, out string text)
    {
        switch (value)
        {
            case null:
                text = string.Empty;
                return null;

            case string s:
                text = s;
                return null;

            case double d:
                text = d.ToString(CultureInfo.InvariantCulture);
                return null;

            case bool b:
                text = b ? "TRUE" : "FALSE";
                return null;

            case ErrorValue error:
                text = string.Empty;
                return error;

            default:
                text = string.Empty;
                return ErrorValue.NotValue;
        }
    }

    /// <summary>
    /// Excel-style equality for the <c>=</c>/<c>&lt;&gt;</c> operators: numbers compare numerically,
    /// strings case-insensitively, and values of different types are never equal (so <c>1="1"</c> is
    /// false). A blank (<c>null</c>) is equal to the "empty" of the other operand: <c>0</c>, <c>""</c>
    /// or <c>false</c>. Callers must propagate <see cref="ErrorValue"/> operands before calling this.
    /// </summary>
    public static bool AreEqual(object? left, object? right)
    {
        if (left is null && right is null)
        {
            return true;
        }

        if (left is null)
        {
            return IsBlankEquivalent(right!);
        }

        if (right is null)
        {
            return IsBlankEquivalent(left);
        }

        return (left, right) switch
        {
            (double l, double r) => l == r,
            (bool l, bool r) => l == r,
            (string l, string r) => string.Equals(l, r, StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    /// <summary>
    /// Excel-style ordering for <c>&lt; &gt; &lt;= &gt;=</c>: numbers sort before text, which sorts before
    /// booleans (FALSE before TRUE); same-type text compares case-insensitively; blank counts as 0.
    /// Callers must propagate <see cref="ErrorValue"/> operands before calling this.
    /// </summary>
    public static int Compare(object? left, object? right)
    {
        var (leftRank, leftNumber, leftText) = Classify(left);
        var (rightRank, rightNumber, rightText) = Classify(right);

        if (leftRank != rightRank)
        {
            return leftRank.CompareTo(rightRank);
        }

        return leftRank == 1
            ? string.Compare(leftText, rightText, StringComparison.OrdinalIgnoreCase)
            : leftNumber.CompareTo(rightNumber);
    }

    private static (int Rank, double Number, string Text) Classify(object? value) =>
        value switch
        {
            null => (0, 0, string.Empty), // blank counts as 0
            double d => (0, d, string.Empty),
            string s => (1, 0, s),
            bool b => (2, b ? 1 : 0, string.Empty),
            _ => (1, 0, value.ToString() ?? string.Empty),
        };

    private static bool IsBlankEquivalent(object value) =>
        value switch
        {
            double d => d == 0,
            string s => s.Length == 0,
            bool b => b == false,
            _ => false,
        };
}
