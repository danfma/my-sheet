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

    /// <summary>
    /// Tokenizes a formula body once so shared-formula slaves can re-parse it with per-cell deltas —
    /// the token list is immutable in practice and reusable across every slave of the group.
    /// </summary>
    internal static List<Token> TokenizeFormulaBody(string body) => Tokenizer.Tokenize(body);

    /// <summary>
    /// Parses a pre-tokenized shared-formula master, shifting every RELATIVE cell reference by the
    /// (row, column) delta — Excel's shared-formula expansion, done on the token stream instead of a
    /// text rewrite + full re-tokenize per slave. '$'-anchored components do not move (the '$' still
    /// lives in the token text; the AST drops it).
    /// </summary>
    internal static Expression ParseSharedFormulaBody(
        List<Token> tokens,
        Sheet sheet,
        int deltaRow,
        int deltaColumn
    )
    {
        return new Parser(tokens, sheet.Name, deltaRow, deltaColumn).ParseFormula();
    }

    private static Expression ParseLiteral(string text)
    {
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
        {
            return new NumberValue(number);
        }

        if (bool.TryParse(text, out var boolean))
        {
            return boolean ? BooleanValue.True : BooleanValue.False;
        }

        return new StringValue(text);
    }
}
