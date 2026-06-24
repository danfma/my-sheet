namespace MySheet.Expressions;

public sealed record ErrorValue(string ErrorCode)
{
    public static readonly ErrorValue NotValue = new("#VALUE!");
}