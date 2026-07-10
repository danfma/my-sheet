using System.Globalization;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Excel-style coercion and comparison over <see cref="ComputedValue"/>, kept internal to the engine (it is
/// not part of the public value surface). The coercion helpers are extension methods, so they read fluently
/// as <c>value.CoerceToNumber(out var n)</c>; they are NOT <c>Try*</c> (they do not return a bool) — they
/// return the <see cref="Error"/> to propagate, or <c>null</c> on success, consumed via
/// <c>if (value.CoerceTo…(out var v) is { } error) return error;</c>.
/// </summary>
internal static class ValueCoercion
{
    /// <summary>Coerces to a number (Excel-style). <c>null</c> on success (value in <paramref name="number"/>),
    /// or the <see cref="Error"/> to propagate on failure.</summary>
    public static Error? CoerceToNumber(this in ComputedValue value, out double number)
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Blank:
                number = 0;
                return null;

            case ComputedValueKind.Number:
                value.TryGetNumber(out number);
                return null;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out var boolean);
                number = boolean ? 1 : 0;
                return null;

            case ComputedValueKind.Text
                when value.TryGetText(out var text)
                    && double.TryParse(
                        text,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var parsed
                    ):
                number = parsed;
                return null;

            case ComputedValueKind.Error:
                value.TryGetError(out var error);
                number = 0;
                return error;

            default:
                number = 0;
                return Error.Value;
        }
    }

    /// <summary>Coerces to a boolean condition (Excel truthiness). <c>null</c> on success, or the <see cref="Error"/> to propagate.</summary>
    public static Error? CoerceToBool(this in ComputedValue value, out bool result)
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Blank:
                result = false;
                return null;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out result);
                return null;

            case ComputedValueKind.Number:
                value.TryGetNumber(out var number);
                result = number != 0;
                return null;

            case ComputedValueKind.Error:
                value.TryGetError(out var error);
                result = false;
                return error;

            default:
                result = false;
                return Error.Value;
        }
    }

    /// <summary>Coerces to text (blank→""; number→invariant; bool→TRUE/FALSE). <c>null</c> on success, or the <see cref="Error"/> to propagate.</summary>
    public static Error? CoerceToText(this in ComputedValue value, out string text)
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Blank:
                text = string.Empty;
                return null;

            case ComputedValueKind.Text:
                value.TryGetText(out text!);
                return null;

            case ComputedValueKind.Number:
                value.TryGetNumber(out var number);
                text = number.ToString(CultureInfo.InvariantCulture);
                return null;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out var boolean);
                text = boolean ? "TRUE" : "FALSE";
                return null;

            case ComputedValueKind.Error:
                value.TryGetError(out var error);
                text = string.Empty;
                return error;

            default:
                text = string.Empty;
                return Error.Value;
        }
    }

    /// <summary>
    /// Excel-style equality for <c>=</c>/<c>&lt;&gt;</c>: numbers compare numerically, strings
    /// case-insensitively, different types are never equal (so <c>1="1"</c> is false); a blank equals the
    /// "empty" of the other operand (<c>0</c>/<c>""</c>/<c>false</c>). Callers propagate errors first.
    /// </summary>
    public static bool AreEqual(in ComputedValue left, in ComputedValue right)
    {
        if (left.Kind == ComputedValueKind.Blank && right.Kind == ComputedValueKind.Blank)
        {
            return true;
        }

        if (left.Kind == ComputedValueKind.Blank)
        {
            return IsBlankEquivalent(right);
        }

        if (right.Kind == ComputedValueKind.Blank)
        {
            return IsBlankEquivalent(left);
        }

        if (left.Kind != right.Kind)
        {
            return false;
        }

        switch (left.Kind)
        {
            case ComputedValueKind.Number:
                left.TryGetNumber(out var leftNumber);
                right.TryGetNumber(out var rightNumber);
                return leftNumber == rightNumber;

            case ComputedValueKind.Boolean:
                left.TryGetBoolean(out var leftBool);
                right.TryGetBoolean(out var rightBool);
                return leftBool == rightBool;

            case ComputedValueKind.Text:
                left.TryGetText(out var leftText);
                right.TryGetText(out var rightText);
                return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);

            default:
                return false;
        }
    }

    /// <summary>
    /// Excel-style ordering for <c>&lt; &gt; &lt;= &gt;=</c>: number &lt; text &lt; boolean (FALSE before
    /// TRUE); same-type text compares case-insensitively; blank counts as 0. Callers propagate errors first.
    /// </summary>
    public static int Compare(in ComputedValue left, in ComputedValue right)
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

    private static (int Rank, double Number, string Text) Classify(in ComputedValue value)
    {
        if (value.TryGetNumber(out var number))
        {
            return (0, number, string.Empty);
        }

        if (value.TryGetText(out var text))
        {
            return (1, 0, text);
        }

        if (value.TryGetBoolean(out var boolean))
        {
            return (2, boolean ? 1 : 0, string.Empty);
        }

        // Blank counts as 0 (rank 0); other kinds (Reference) are conservatively treated as text.
        return value.Kind == ComputedValueKind.Blank ? (0, 0, string.Empty) : (1, 0, string.Empty);
    }

    private static bool IsBlankEquivalent(in ComputedValue value)
    {
        if (value.TryGetNumber(out var number))
        {
            return number == 0;
        }

        if (value.TryGetText(out var text))
        {
            return text.Length == 0;
        }

        if (value.TryGetBoolean(out var boolean))
        {
            return !boolean;
        }

        return false;
    }
}
