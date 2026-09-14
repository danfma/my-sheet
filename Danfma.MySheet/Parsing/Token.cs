namespace Danfma.MySheet.Parsing;

internal enum TokenType
{
    Number,
    String,
    Identifier,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Equal,
    NotEqual,
    Less,
    Greater,
    LessEqual,
    GreaterEqual,
    Ampersand,
    Percent,
    Comma,
    Colon,
    LParen,
    RParen,
    Bang,
    DeletedReferenceSpill,

    // The whole `[...]` suffix of a structured (table) reference, Text = the raw bracketed text INCLUDING
    // the brackets, undecoded.
    BracketedSpecifier,

    // One of the 7 classic Excel error literals (#NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A) —
    // the spelling Excel itself writes into formula text for a broken reference.
    // Text is always the CANONICAL upper-case spelling, regardless of the input's case (Aspose.Cells
    // 26.7.0/26.6.0 both accept any case and normalize the same way on round trip).
    Error,
    EndOfInput,
}

internal readonly record struct Token(TokenType Type, string Text, int Position);
