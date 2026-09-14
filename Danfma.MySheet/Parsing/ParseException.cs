namespace Danfma.MySheet.Parsing;

/// <summary>
/// The category of a <see cref="ParseException"/>, so integrators can report or handle syntax failures
/// by shape instead of by parsing the message.
/// </summary>
public enum ParseErrorKind
{
    /// <summary>
    /// No structured category: the exception was raised through the legacy
    /// <c>ParseException(string message, int position)</c> constructor, which predates <see cref="ParseErrorKind"/>.
    /// The parser itself always sets a specific kind.
    /// </summary>
    Unspecified,

    /// <summary>A character the formula grammar never uses (<c>=1 # 2</c>).</summary>
    UnexpectedCharacter,

    /// <summary>A string literal whose closing <c>"</c> is missing.</summary>
    UnterminatedString,

    /// <summary>A quoted sheet name whose closing <c>'</c> is missing.</summary>
    UnterminatedQuotedName,

    /// <summary>A token that cannot start or continue an expression at that point (<c>=*2</c>, <c>=1 2</c>).</summary>
    UnexpectedToken,

    /// <summary>A specific token was required — typically a closing <c>)</c> or a <c>,</c> — and something else was found.</summary>
    ExpectedToken,

    /// <summary>
    /// The text after a sheet qualifier's <c>!</c> is not a cell, column, row or range. Also raised for a
    /// <c>:</c> range endpoint that is an error literal other than <c>#REF!</c> (<c>A1:#N/A</c>) — Excel's
    /// own broken-reference spelling is the only error accepted there.
    /// </summary>
    ExpectedCellReference,

    /// <summary>A built-in function was called with an argument count it does not accept (<c>=ROUND(1)</c>).</summary>
    InvalidArgumentCount,

    /// <summary>The formula nests deeper than the parser's recursion limit.</summary>
    NestingTooDeep,

    /// <summary>A structured (table) reference whose closing <c>]</c> is missing (<c>=Tabela1[Valor</c>).</summary>
    UnterminatedBracketedReference,

    /// <summary>
    /// Balanced brackets whose content Excel's own grammar rejects — an unknown specifier
    /// (<c>Tabela1[#Bogus]</c>), a dangling <c>'</c> escape, a multi-column item list
    /// (<c>Tabela1[[A],[B]]</c>). The file is malformed, as opposed to
    /// <see cref="UnsupportedStructuredReference"/>.
    /// </summary>
    InvalidStructuredReference,

    /// <summary>
    /// Valid Excel that MySheet does not model — a scope decision, not parity: the current-row forms
    /// <c>Tabela1[@Valor]</c> and <c>Tabela1[[#This Row],[Valor]]</c>, the implicit-table form <c>[Valor]</c>,
    /// a column span <c>Tabela1[[A]:[C]]</c>, the external-workbook form <c>[1]Sheet1!A1</c>, and a
    /// sheet-qualified <c>Data!Tabela1[Valor]</c>. When it is uncertain whether Excel accepts a shape it is
    /// classified here, never as <see cref="InvalidStructuredReference"/>.
    /// </summary>
    UnsupportedStructuredReference,
}

/// <summary>
/// Raised for SYNTAX errors while parsing a formula — the formula cannot be turned into an expression tree at
/// all, which is a different category from a formula that parses but fails to calculate. Semantic failures
/// (an unknown function, a missing sheet, a bad range) never throw: they evaluate to <c>ErrorValue</c>s such as
/// <c>#NAME?</c> or <c>#REF!</c>. So catching this exception is, by itself, the "syntax vs. semantic"
/// distinction; <see cref="Kind"/>, <see cref="Token"/> and <see cref="Position"/> then say what and where.
/// </summary>
public sealed class ParseException(ParseErrorKind kind, string message, int position, string token)
    : Exception($"{message} (at position {position}).")
{
    /// <summary>
    /// Compatibility overload (the pre-3.16 shape): the same message and position, with
    /// <see cref="Kind"/> = <see cref="ParseErrorKind.Unspecified"/> and an empty <see cref="Token"/>.
    /// </summary>
    public ParseException(string message, int position)
        : this(ParseErrorKind.Unspecified, message, position, string.Empty) { }

    /// <summary>What kind of syntax error this is.</summary>
    public ParseErrorKind Kind { get; } = kind;

    /// <summary>
    /// The 0-based offset of the offending token in the formula BODY — the text after the leading <c>=</c>
    /// (so <c>=1 2</c> reports position 2 for the <c>2</c>). Formulas parsed via
    /// <see cref="ExpressionParser.ParseFormulaBody"/> have no <c>=</c> and the offset applies to the string
    /// as given.
    /// </summary>
    public int Position { get; } = position;

    /// <summary>
    /// The offending token's text as it appeared in the formula (the unexpected character, the token found
    /// where another was expected, the function name whose arity is wrong, the unterminated literal's
    /// remainder). Empty when the parser ran out of input, e.g. an unclosed parenthesis.
    /// </summary>
    public string Token { get; } = token;
}
