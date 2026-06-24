using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Mid(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(workbook), out var text) is { } textError)
        {
            return textError;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(workbook), out var start) is { } startError)
        {
            return startError;
        }

        if (ValueCoercion.TryToNumber(Arguments[2].Compute(workbook), out var count) is { } countError)
        {
            return countError;
        }

        if (start < 1 || count < 0)
        {
            return ErrorValue.NotValue;
        }

        if (start > text.Length)
        {
            return string.Empty;
        }

        var available = text.Length - ((int)start - 1);

        return text.Substring((int)start - 1, Math.Min((int)count, available));
    }
}
