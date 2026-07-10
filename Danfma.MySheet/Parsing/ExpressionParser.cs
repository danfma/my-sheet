using System.Globalization;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Parsing;

public static class ExpressionParser
{
    /// <summary>
    /// Parses a cell entry into an <see cref="Expression"/>. Entries starting with '=' are parsed as
    /// formulas; anything else is treated as a literal value (number, boolean or text). Syntax errors
    /// throw <see cref="ParseException"/>; semantic errors (unknown function, etc.) become
    /// <c>ErrorValue</c> nodes.
    /// </summary>
    public static Expression Parse(string expression, Sheet sheet)
    {
        ArgumentNullException.ThrowIfNull(expression);

        if (expression.Length == 0)
        {
            return BlankValue.Instance;
        }

        if (expression[0] == '=')
        {
            var tokens = Tokenizer.Tokenize(expression[1..]);

            return new Parser(tokens, sheet.Name).ParseFormula();
        }

        return ParseLiteral(expression);
    }

    /// <summary>
    /// Parses a formula BODY — the text after the leading '='. Callers that already hold the body
    /// (the .xlsx loader, INDIRECT) skip the two per-formula string copies <see cref="Parse"/> would
    /// cost them: the <c>"=" + body</c> concatenation and the slice back. Same semantics as
    /// <c>Parse("=" + body, sheet)</c>.
    /// </summary>
    public static Expression ParseFormulaBody(string body, Sheet sheet)
    {
        ArgumentNullException.ThrowIfNull(body);

        return new Parser(Tokenizer.Tokenize(body), sheet.Name).ParseFormula();
    }

    private static Expression ParseLiteral(string text)
    {
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
        {
            return new NumberValue(number);
        }

        if (bool.TryParse(text, out var boolean))
        {
            return new BooleanValue(boolean);
        }

        return new StringValue(text);
    }
}
