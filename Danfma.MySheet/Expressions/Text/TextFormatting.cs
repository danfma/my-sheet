using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

// Onda 2 — formatação e extração de texto: FIXED, DOLLAR (invariant: '.' decimal, ',' milhar, '$'
// — contrato locale-invariant do §A7 do roadmap), NUMBERVALUE, TEXTBEFORE/TEXTAFTER e VALUETOTEXT.

[MemoryPackable]
public sealed partial record Fixed(Expression[] Arguments) : Function
{
    // FIXED(number, [decimals = 2], [no_commas = FALSE]) — returns TEXT. Negative decimals round
    // to the left of the decimal point; the doc caps decimals at 127.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var decimals = 2.0;

        if (
            Arguments.Length >= 2
            && Arguments[1] is not BlankValue
            && Arguments[1].Evaluate(context).CoerceToNumber(out decimals) is { } decimalsError
        )
        {
            return ComputedValue.Error(decimalsError);
        }

        var noCommas = false;

        if (
            Arguments.Length >= 3
            && Arguments[2].Evaluate(context).CoerceToBool(out noCommas) is { } commasError
        )
        {
            return ComputedValue.Error(commasError);
        }

        var digits = (int)Math.Truncate(decimals);

        if (digits > 127)
        {
            return ComputedValue.Error(Error.Value);
        }

        var rounded = NumberFormatting.RoundToDigits(number, digits);
        var displayDigits = Math.Max(digits, 0);

        return ComputedValue.Text(
            rounded.ToString(
                (noCommas ? "F" : "N") + displayDigits.ToString(CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture
            )
        );
    }
}

[MemoryPackable]
public sealed partial record Dollar(Expression[] Arguments) : Function
{
    // DOLLAR(number, [decimals = 2]) — returns TEXT in the documented $#,##0.00_);($#,##0.00)
    // shape: "$1,234.57" and "($1,200)" for negatives. Invariant '$', per the §A7 contract.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var decimals = 2.0;

        if (
            Arguments.Length >= 2
            && Arguments[1] is not BlankValue
            && Arguments[1].Evaluate(context).CoerceToNumber(out decimals) is { } decimalsError
        )
        {
            return ComputedValue.Error(decimalsError);
        }

        var digits = (int)Math.Truncate(decimals);

        if (digits > 127)
        {
            return ComputedValue.Error(Error.Value);
        }

        var rounded = NumberFormatting.RoundToDigits(number, digits);
        var magnitude = Math.Abs(rounded)
            .ToString(
                "N" + Math.Max(digits, 0).ToString(CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture
            );

        return ComputedValue.Text(rounded < 0 ? $"(${magnitude})" : $"${magnitude}");
    }
}

[MemoryPackable]
public sealed partial record NumberValueFunction(Expression[] Arguments) : Function
{
    // NUMBERVALUE(text, [decimal_separator = '.'], [group_separator = ',']) — locale-independent
    // text-to-number: spaces are ignored anywhere; trailing % signs divide by 100 each; a decimal
    // separator used twice, or a group separator after the decimal separator -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        var decimalSeparator = '.';
        var groupSeparator = ',';

        if (Arguments.Length >= 2 && Arguments[1] is not BlankValue)
        {
            if (Arguments[1].Evaluate(context).CoerceToText(out var separator) is { } error)
            {
                return ComputedValue.Error(error);
            }

            if (separator.Length == 0)
            {
                return ComputedValue.Error(Error.Value);
            }

            decimalSeparator = separator[0]; // only the first character is used (per the docs)
        }

        if (Arguments.Length >= 3 && Arguments[2] is not BlankValue)
        {
            if (Arguments[2].Evaluate(context).CoerceToText(out var separator) is { } error)
            {
                return ComputedValue.Error(error);
            }

            if (separator.Length == 0)
            {
                return ComputedValue.Error(Error.Value);
            }

            groupSeparator = separator[0];
        }

        if (decimalSeparator == groupSeparator)
        {
            return ComputedValue.Error(Error.Value);
        }

        var compact = string.Concat(text.Where(c => !char.IsWhiteSpace(c)));

        if (compact.Length == 0)
        {
            return ComputedValue.Number(0); // empty text -> 0 (per the docs)
        }

        var percents = 0;

        while (compact.EndsWith('%'))
        {
            percents++;
            compact = compact[..^1];
        }

        var decimalIndex = compact.IndexOf(decimalSeparator);

        if (decimalIndex >= 0 && compact.IndexOf(decimalSeparator, decimalIndex + 1) >= 0)
        {
            return ComputedValue.Error(Error.Value); // decimal separator used more than once
        }

        if (decimalIndex >= 0 && compact.IndexOf(groupSeparator, decimalIndex + 1) >= 0)
        {
            return ComputedValue.Error(Error.Value); // group separator after the decimal separator
        }

        var normalized = string.Concat(
            compact
                .Where(c => c != groupSeparator)
                .Select(c => c == decimalSeparator ? '.' : c)
        );

        if (
            !double.TryParse(
                normalized,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var number
            )
        )
        {
            return ComputedValue.Error(Error.Value);
        }

        return ComputedValue.Number(number / Math.Pow(100, percents));
    }
}

[MemoryPackable]
public sealed partial record TextBefore(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        DelimiterSplit.Evaluate(Arguments, context, after: false);
}

[MemoryPackable]
public sealed partial record TextAfter(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        DelimiterSplit.Evaluate(Arguments, context, after: true);
}

// TEXTBEFORE/TEXTAFTER(text, delimiter, [instance_num = 1], [match_mode = 0], [match_end = 0],
// [if_not_found]) — shared engine. A negative instance_num counts occurrences from the end;
// match_mode 1 is case-insensitive; match_end 1 treats the end of the text (start, when counting
// from the end) as one extra delimiter; a miss returns if_not_found when given, otherwise #N/A.
file static class DelimiterSplit
{
    public static ComputedValue Evaluate(Expression[] arguments, EvaluationContext context, bool after)
    {
        if (arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (arguments[1].Evaluate(context).CoerceToText(out var delimiter) is { } delimiterError)
        {
            return ComputedValue.Error(delimiterError);
        }

        var instance = 1.0;

        if (
            arguments.Length >= 3
            && arguments[2] is not BlankValue
            && arguments[2].Evaluate(context).CoerceToNumber(out instance) is { } instanceError
        )
        {
            return ComputedValue.Error(instanceError);
        }

        var matchMode = 0.0;

        if (
            arguments.Length >= 4
            && arguments[3] is not BlankValue
            && arguments[3].Evaluate(context).CoerceToNumber(out matchMode) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        var matchEnd = 0.0;

        if (
            arguments.Length >= 5
            && arguments[4] is not BlankValue
            && arguments[4].Evaluate(context).CoerceToNumber(out matchEnd) is { } endError
        )
        {
            return ComputedValue.Error(endError);
        }

        if (matchMode is not (0 or 1) || matchEnd is not (0 or 1))
        {
            return ComputedValue.Error(Error.Value);
        }

        if (text.Length == 0)
        {
            return ComputedValue.Text(string.Empty); // empty text -> empty text (per the docs)
        }

        var n = (int)Math.Truncate(instance);

        if (n == 0 || Math.Abs(n) > text.Length)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (delimiter.Length == 0)
        {
            // An empty delimiter matches immediately: at the front for a positive instance_num,
            // at the very end for a negative one (per the docs).
            return ComputedValue.Text(after == n > 0 ? text : string.Empty);
        }

        var comparison = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var positions = new List<int>();
        var index = 0;

        while ((index = text.IndexOf(delimiter, index, comparison)) >= 0)
        {
            positions.Add(index);
            index += delimiter.Length;
        }

        int start, end;
        var occurrence = Math.Abs(n);

        if (occurrence <= positions.Count)
        {
            start = n > 0 ? positions[occurrence - 1] : positions[^occurrence];
            end = start + delimiter.Length;
        }
        else if (matchEnd == 1 && occurrence == positions.Count + 1)
        {
            // The virtual delimiter: the end of the text (or its start, counting from the end).
            start = end = n > 0 ? text.Length : 0;
        }
        else
        {
            return arguments.Length >= 6 && arguments[5] is not BlankValue
                ? arguments[5].Evaluate(context)
                : ComputedValue.Error(Error.NA);
        }

        return ComputedValue.Text(after ? text[end..] : text[..start]);
    }
}

[MemoryPackable]
public sealed partial record ValueToText(Expression[] Arguments) : Function
{
    // VALUETOTEXT(value, [format = 0]) — 0 renders like the cell (concise); 1 (strict) wraps text
    // in quotes (inner quotes doubled) but leaves Booleans, Numbers and Errors bare. Any other
    // format -> #VALUE!. Errors become their display text instead of propagating.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var format = 0.0;

        if (
            Arguments.Length >= 2
            && Arguments[1] is not BlankValue
            && Arguments[1].Evaluate(context).CoerceToNumber(out format) is { } formatError
        )
        {
            return ComputedValue.Error(formatError);
        }

        if (format is not (0 or 1))
        {
            return ComputedValue.Error(Error.Value);
        }

        var value = Arguments[0].Evaluate(context);

        if (value.TryGetError(out var error))
        {
            return ComputedValue.Text(error.Display);
        }

        if (value.TryGetText(out var text))
        {
            return ComputedValue.Text(
                format == 1 ? "\"" + text.Replace("\"", "\"\"") + "\"" : text
            );
        }

        return value.CoerceToText(out var rendered) is { } renderError
            ? ComputedValue.Error(renderError)
            : ComputedValue.Text(rendered);
    }
}

// FIXED/DOLLAR shared rounding: Excel's ROUND (midpoint away from zero) at a digit position that
// may be negative (left of the decimal point).
file static class NumberFormatting
{
    public static double RoundToDigits(double number, int digits)
    {
        if (digits >= 0)
        {
            // Digits beyond double precision change nothing — formatting pads with zeros.
            return digits > 15 ? number : Math.Round(number, digits, MidpointRounding.AwayFromZero);
        }

        var factor = Math.Pow(10, -digits);

        return Math.Round(number / factor, MidpointRounding.AwayFromZero) * factor;
    }
}
