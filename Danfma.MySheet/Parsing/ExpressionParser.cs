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

    /// <summary>
    /// G3 spike (node-delta shared formulas): parses a shared-formula group's MASTER once in the Parser's
    /// ANCHORED mode — every relative cell/range reference becomes an <see cref="AnchoredCellReference"/>/
    /// <see cref="AnchoredRangeReference"/> that keeps its ($-anchor, column, row) components instead of
    /// being shifted into a per-slave string id. The resulting tree is meant to be shared by every slave in
    /// the group via <see cref="SharedFormulaSlave"/> (each slave supplies only its own (row, column) delta
    /// at evaluation time) — see <c>WorksheetStreamLoader.ExpandSlave</c> for the caller, which additionally
    /// verifies (via <c>AnchoredFormulaSupport.IsFullyAnchored</c>) that the tree contains no node this mode
    /// cannot represent (an open range, a union, a reference-returning endpoint) before trusting it for the
    /// whole group; a <see cref="ParseException"/> from a shape neither this method nor that check handles
    /// gracefully (e.g. a chained cross-sheet range endpoint) is likewise treated by the caller as "fall back
    /// to the legacy per-slave token-delta expansion" — this method itself does not need to special-case
    /// every such shape, since the caller's shape check and its own try/catch are the safety net.
    /// </summary>
    internal static Expression ParseAnchoredMasterBody(List<Token> tokens, Sheet sheet)
    {
        return new Parser(tokens, sheet.Name, anchored: true).ParseFormula();
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
