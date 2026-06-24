namespace MySheet.Parsing;

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
    EndOfInput,
}

internal readonly record struct Token(TokenType Type, string Text, int Position);
