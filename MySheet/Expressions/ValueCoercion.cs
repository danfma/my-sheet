using System.Globalization;

namespace MySheet.Expressions;

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
                when double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed):
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
}
