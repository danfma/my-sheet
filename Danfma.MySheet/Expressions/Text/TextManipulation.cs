using System.Text;
using System.Text.RegularExpressions;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

// Onda 2 — manipulação de texto: RIGHT, FIND (case-sensitive, sem wildcards), SEARCH
// (case-insensitive, com wildcards ? * ~, busca POSICIONAL), REPLACE, SUBSTITUTE, REPT, PROPER,
// EXACT, CHAR, CODE, UNICHAR, UNICODE e CLEAN. Contrato locale-invariant (§A7 do roadmap):
// comparações ordinais, maiúsculas/minúsculas invariant. CHAR/CODE usam o code point Unicode
// (Latin-1 para 1-255), não a página ANSI do Windows — limitação documentada no reference.

[MemoryPackable]
public sealed partial record Right(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        var count = 1.0;

        if (
            Arguments.Length == 2
            && Arguments[1].Evaluate(context).CoerceToNumber(out count) is { } countError
        )
        {
            return ComputedValue.Error(countError);
        }

        if (count < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        return ComputedValue.Text(text[^Math.Min((int)count, text.Length)..]);
    }
}

[MemoryPackable]
public sealed partial record Find(Expression[] Arguments) : Function
{
    // Case-sensitive ordinal search, no wildcards. Empty find_text matches at start_num.
    // Not found, start_num < 1, or start_num beyond within_text -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var findText) is { } findError)
        {
            return ComputedValue.Error(findError);
        }

        if (Arguments[1].Evaluate(context).CoerceToText(out var withinText) is { } withinError)
        {
            return ComputedValue.Error(withinError);
        }

        var start = 1.0;

        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out start) is { } startError
        )
        {
            return ComputedValue.Error(startError);
        }

        var startIndex = (int)start - 1;

        if (startIndex < 0 || startIndex >= withinText.Length + (findText.Length == 0 ? 1 : 0))
        {
            return ComputedValue.Error(Error.Value);
        }

        var index = withinText.IndexOf(findText, startIndex, StringComparison.Ordinal);

        return index < 0 ? ComputedValue.Error(Error.Value) : ComputedValue.Number(index + 1);
    }
}

[MemoryPackable]
public sealed partial record Search(Expression[] Arguments) : Function
{
    // Case-insensitive POSITIONAL search with the Excel wildcards: '?' matches any single
    // character, '*' any sequence, and '~' escapes the next character. The result is the 1-based
    // position where the pattern starts inside within_text (leftmost match), not a full match.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var findText) is { } findError)
        {
            return ComputedValue.Error(findError);
        }

        if (Arguments[1].Evaluate(context).CoerceToText(out var withinText) is { } withinError)
        {
            return ComputedValue.Error(withinError);
        }

        var start = 1.0;

        if (
            Arguments.Length == 3
            && Arguments[2].Evaluate(context).CoerceToNumber(out start) is { } startError
        )
        {
            return ComputedValue.Error(startError);
        }

        var startIndex = (int)start - 1;

        if (startIndex < 0 || startIndex >= withinText.Length + (findText.Length == 0 ? 1 : 0))
        {
            return ComputedValue.Error(Error.Value);
        }

        if (findText.Length == 0)
        {
            return ComputedValue.Number(startIndex + 1);
        }

        try
        {
            var match = WildcardRegex(findText).Match(withinText, startIndex);

            return match.Success
                ? ComputedValue.Number(match.Index + 1)
                : ComputedValue.Error(Error.Value);
        }
        catch (RegexMatchTimeoutException)
        {
            return ComputedValue.Error(Error.Value);
        }
    }

    // The Excel wildcard pattern as an unanchored regex: matching it leftmost-first is exactly
    // SEARCH's positional semantics. The pattern only produces literals, '.' and '.*', so the
    // 1-second timeout is purely defensive.
    private static Regex WildcardRegex(string pattern)
    {
        var builder = new StringBuilder();

        for (var i = 0; i < pattern.Length; i++)
        {
            switch (pattern[i])
            {
                case '~' when i + 1 < pattern.Length:
                    builder.Append(Regex.Escape(pattern[++i].ToString()));
                    break;

                case '*':
                    builder.Append(".*");
                    break;

                case '?':
                    builder.Append('.');
                    break;

                case var c:
                    builder.Append(Regex.Escape(c.ToString()));
                    break;
            }
        }

        return new Regex(
            builder.ToString(),
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline,
            TimeSpan.FromSeconds(1)
        );
    }
}

[MemoryPackable]
public sealed partial record Replace(Expression[] Arguments) : Function
{
    // REPLACE(old_text, start_num, num_chars, new_text) — 1-based positions; start_num < 1 or
    // num_chars < 0 -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var oldText) is { } oldError)
        {
            return ComputedValue.Error(oldError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var start) is { } startError)
        {
            return ComputedValue.Error(startError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var count) is { } countError)
        {
            return ComputedValue.Error(countError);
        }

        if (Arguments[3].Evaluate(context).CoerceToText(out var newText) is { } newError)
        {
            return ComputedValue.Error(newError);
        }

        if (start < 1 || count < 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        var prefixLength = Math.Min((int)start - 1, oldText.Length);
        var suffixStart = Math.Min(prefixLength + (int)count, oldText.Length);

        return ComputedValue.Text(oldText[..prefixLength] + newText + oldText[suffixStart..]);
    }
}

[MemoryPackable]
public sealed partial record Substitute(Expression[] Arguments) : Function
{
    // SUBSTITUTE(text, old_text, new_text, [instance_num]) — case-sensitive ordinal; without
    // instance_num every occurrence is replaced; instance_num is 1-based (< 1 -> #VALUE!).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (Arguments[1].Evaluate(context).CoerceToText(out var oldText) is { } oldError)
        {
            return ComputedValue.Error(oldError);
        }

        if (Arguments[2].Evaluate(context).CoerceToText(out var newText) is { } newError)
        {
            return ComputedValue.Error(newError);
        }

        if (oldText.Length == 0)
        {
            return ComputedValue.Text(text);
        }

        if (Arguments.Length < 4)
        {
            return ComputedValue.Text(text.Replace(oldText, newText, StringComparison.Ordinal));
        }

        if (Arguments[3].Evaluate(context).CoerceToNumber(out var instance) is { } instanceError)
        {
            return ComputedValue.Error(instanceError);
        }

        var target = (int)instance;

        if (target < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        var index = -1;

        for (var found = 0; found < target; found++)
        {
            index = text.IndexOf(oldText, index + 1, StringComparison.Ordinal);

            if (index < 0)
            {
                return ComputedValue.Text(text); // fewer occurrences than instance_num
            }
        }

        return ComputedValue.Text(text[..index] + newText + text[(index + oldText.Length)..]);
    }
}

[MemoryPackable]
public sealed partial record Rept(Expression[] Arguments) : Function
{
    // The count is truncated; 0 -> ""; negative -> #VALUE!; a result longer than Excel's 32,767
    // character cell limit -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var count) is { } countError)
        {
            return ComputedValue.Error(countError);
        }

        var times = Math.Truncate(count);

        if (times < 0 || text.Length * times > 32767)
        {
            return ComputedValue.Error(Error.Value);
        }

        return ComputedValue.Text(string.Concat(Enumerable.Repeat(text, (int)times)));
    }
}

[MemoryPackable]
public sealed partial record Proper(Expression[] Arguments) : Function
{
    // Upper-cases every letter that follows a non-letter (including the first), lower-cases the
    // rest: "2-way street" -> "2-Way Street", "76BudGet" -> "76Budget".
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var builder = new StringBuilder(text.Length);
        var previousIsLetter = false;

        foreach (var c in text)
        {
            if (char.IsLetter(c))
            {
                builder.Append(
                    previousIsLetter ? char.ToLowerInvariant(c) : char.ToUpperInvariant(c)
                );
                previousIsLetter = true;
            }
            else
            {
                builder.Append(c);
                previousIsLetter = false;
            }
        }

        return ComputedValue.Text(builder.ToString());
    }
}

[MemoryPackable]
public sealed partial record Exact(Expression[] Arguments) : Function
{
    // Case-sensitive (ordinal) comparison; non-text values compare by their text form.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var left) is { } leftError)
        {
            return ComputedValue.Error(leftError);
        }

        if (Arguments[1].Evaluate(context).CoerceToText(out var right) is { } rightError)
        {
            return ComputedValue.Error(rightError);
        }

        return ComputedValue.Boolean(string.Equals(left, right, StringComparison.Ordinal));
    }
}

[MemoryPackable]
public sealed partial record CharFunction(Expression[] Arguments) : Function
{
    // CHAR(number) — 1-255 (truncated), outside -> #VALUE!. The engine maps the code as a Unicode
    // code point (Latin-1), not the Windows ANSI code page (locale-invariant contract).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var code = (int)Math.Truncate(number);

        return code is < 1 or > 255
            ? ComputedValue.Error(Error.Value)
            : ComputedValue.Text(((char)code).ToString());
    }
}

[MemoryPackable]
public sealed partial record Code(Expression[] Arguments) : Function
{
    // CODE(text) — the code of the first character; empty text -> #VALUE!. Returns the Unicode
    // code point (a surrogate pair counts as one character), same scale as UNICODE.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return FirstCodePoint(text);
    }

    internal static ComputedValue FirstCodePoint(string text)
    {
        if (text.Length == 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (char.IsHighSurrogate(text[0]) && text.Length > 1 && char.IsLowSurrogate(text[1]))
        {
            return ComputedValue.Number(char.ConvertToUtf32(text[0], text[1]));
        }

        // An unpaired surrogate is not a valid code point (per the UNICODE docs -> #VALUE!).
        return char.IsSurrogate(text[0])
            ? ComputedValue.Error(Error.Value)
            : ComputedValue.Number(text[0]);
    }
}

[MemoryPackable]
public sealed partial record UniChar(Expression[] Arguments) : Function
{
    // UNICHAR(number) — full Unicode. 0/negative/above 10FFFF -> #VALUE!; a surrogate code point
    // (D800-DFFF, not a valid scalar) -> #N/A, per the Excel docs.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var code = (int)Math.Truncate(number);

        if (code is < 1 or > 0x10FFFF)
        {
            return ComputedValue.Error(Error.Value);
        }

        return code is >= 0xD800 and <= 0xDFFF
            ? ComputedValue.Error(Error.NA)
            : ComputedValue.Text(char.ConvertFromUtf32(code));
    }
}

[MemoryPackable]
public sealed partial record Unicode(Expression[] Arguments) : Function
{
    // UNICODE(text) — the code point of the first character; empty or invalid -> #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return Code.FirstCodePoint(text);
    }
}

[MemoryPackable]
public sealed partial record Clean(Expression[] Arguments) : Function
{
    // Removes the 7-bit ASCII control characters (codes 0-31) only — 127 etc. stay, per the docs.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var builder = new StringBuilder(text.Length);

        foreach (var c in text)
        {
            if (c >= 32)
            {
                builder.Append(c);
            }
        }

        return ComputedValue.Text(builder.ToString());
    }
}
