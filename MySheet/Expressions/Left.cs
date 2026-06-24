using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Left(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(workbook), out var text) is { } textError)
        {
            return textError;
        }

        var count = 1.0;

        if (Arguments.Length == 2 &&
            ValueCoercion.TryToNumber(Arguments[1].Compute(workbook), out count) is { } countError)
        {
            return countError;
        }

        if (count < 0)
        {
            return ErrorValue.NotValue;
        }

        return text[..Math.Min((int)count, text.Length)];
    }
}
