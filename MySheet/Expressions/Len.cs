using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Len(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(workbook), out var text) is { } error)
        {
            return error;
        }

        return (double)text.Length;
    }
}
