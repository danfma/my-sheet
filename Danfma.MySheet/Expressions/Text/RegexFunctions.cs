using System.Text.RegularExpressions;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

// Onda 2 — regex (REGEXTEST, REGEXEXTRACT, REGEXREPLACE) via System.Text.RegularExpressions com
// timeout defensivo de 1s. O Excel especifica o flavor PCRE2; o .NET cobre o subconjunto usual
// (classes, quantificadores, âncoras, grupos, $n) — limitação documentada no function reference.
// case_sensitivity: 0 = sensitive (default), 1 = insensitive; outros valores -> #VALUE!.

[MemoryPackable]
public sealed partial record RegexTest(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (ExcelRegex.Create(Arguments, 2, context, out var regex) is { } regexError)
        {
            return ComputedValue.Error(regexError);
        }

        try
        {
            return ComputedValue.Boolean(regex.IsMatch(text));
        }
        catch (RegexMatchTimeoutException)
        {
            return ComputedValue.Error(Error.Value);
        }
    }
}

[MemoryPackable]
public sealed partial record RegexExtract(Expression[] Arguments) : Function
{
    // REGEXEXTRACT(text, pattern, [return_mode = 0], [case_sensitivity = 0]) — only the scalar
    // return_mode 0 (first match) is supported; modes 1 (all matches) and 2 (capturing groups)
    // return arrays and stay #VALUE! until the arrays phase (F2). No match -> #N/A.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        var returnMode = 0.0;

        if (
            Arguments.Length >= 3
            && Arguments[2] is not BlankValue
            && Arguments[2].Evaluate(context).CoerceToNumber(out returnMode) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        if (returnMode != 0)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (ExcelRegex.Create(Arguments, 3, context, out var regex) is { } regexError)
        {
            return ComputedValue.Error(regexError);
        }

        try
        {
            var match = regex.Match(text);

            return match.Success ? ComputedValue.Text(match.Value) : ComputedValue.Error(Error.NA);
        }
        catch (RegexMatchTimeoutException)
        {
            return ComputedValue.Error(Error.Value);
        }
    }
}

[MemoryPackable]
public sealed partial record RegexReplace(Expression[] Arguments) : Function
{
    // REGEXREPLACE(text, pattern, replacement, [occurrence = 0], [case_sensitivity = 0]) —
    // occurrence 0 replaces every instance; a positive n only the nth; a negative n counts from
    // the end. Capturing groups are referenced as $n in the replacement. When the requested
    // occurrence does not exist the text is returned unchanged (undocumented in the Excel page;
    // engine decision).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } textError)
        {
            return ComputedValue.Error(textError);
        }

        if (
            Arguments[2].Evaluate(context).CoerceToText(out var replacement) is { } replacementError
        )
        {
            return ComputedValue.Error(replacementError);
        }

        var occurrence = 0.0;

        if (
            Arguments.Length >= 4
            && Arguments[3] is not BlankValue
            && Arguments[3].Evaluate(context).CoerceToNumber(out occurrence) is { } occurrenceError
        )
        {
            return ComputedValue.Error(occurrenceError);
        }

        if (ExcelRegex.Create(Arguments, 4, context, out var regex) is { } regexError)
        {
            return ComputedValue.Error(regexError);
        }

        try
        {
            var target = (int)Math.Truncate(occurrence);

            if (target == 0)
            {
                return ComputedValue.Text(regex.Replace(text, replacement));
            }

            var matches = regex.Matches(text);
            var index = target > 0 ? target - 1 : matches.Count + target;

            if (index < 0 || index >= matches.Count)
            {
                return ComputedValue.Text(text);
            }

            var match = matches[index];

            return ComputedValue.Text(
                text[..match.Index]
                    + match.Result(replacement)
                    + text[(match.Index + match.Length)..]
            );
        }
        catch (RegexMatchTimeoutException)
        {
            return ComputedValue.Error(Error.Value);
        }
    }
}

file static class ExcelRegex
{
    /// <summary>
    /// Builds the regex from <c>Arguments[1]</c> (pattern) and the optional case-sensitivity flag
    /// at <paramref name="caseArgumentIndex"/>. Returns the <see cref="Error"/> to propagate
    /// (invalid pattern or flag -> #VALUE!), or <c>null</c> with the regex ready.
    /// </summary>
    public static Error? Create(
        Expression[] arguments,
        int caseArgumentIndex,
        EvaluationContext context,
        out Regex regex
    )
    {
        regex = null!;

        if (arguments[1].Evaluate(context).CoerceToText(out var pattern) is { } patternError)
        {
            return patternError;
        }

        var caseSensitivity = 0.0;

        if (
            arguments.Length > caseArgumentIndex
            && arguments[caseArgumentIndex] is not BlankValue
            && arguments[caseArgumentIndex].Evaluate(context).CoerceToNumber(out caseSensitivity)
                is { } caseError
        )
        {
            return caseError;
        }

        if (caseSensitivity is not (0 or 1))
        {
            return Error.Value;
        }

        var options =
            RegexOptions.CultureInvariant
            | (caseSensitivity == 1 ? RegexOptions.IgnoreCase : RegexOptions.None);

        try
        {
            regex = new Regex(pattern, options, TimeSpan.FromSeconds(1));
            return null;
        }
        catch (ArgumentException)
        {
            return Error.Value; // invalid pattern
        }
    }
}
