using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Upper(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(workbook), out var text) is { } error)
        {
            return error;
        }

        return text.ToUpperInvariant();
    }
}
