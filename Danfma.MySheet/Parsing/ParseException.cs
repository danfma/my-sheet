namespace Danfma.MySheet.Parsing;

/// <summary>
/// Raised for syntax errors while parsing a formula. Semantic errors (unknown function, etc.)
/// are represented as <c>ErrorValue</c> nodes instead and never throw.
/// </summary>
public sealed class ParseException(string message, int position)
    : Exception($"{message} (at position {position}).")
{
    public int Position { get; } = position;
}
