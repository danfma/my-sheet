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

    // The whole `[...]` suffix of a structured (table) reference, Text = the raw bracketed text INCLUDING
    // the brackets, undecoded.
    BracketedSpecifier,
    EndOfInput,
}

internal readonly record struct Token(TokenType Type, string Text, int Position);
