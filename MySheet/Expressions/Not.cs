using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Not(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToBool(Arguments[0].Compute(workbook), out var value) is { } error)
        {
            return error;
        }

        return !value;
    }
}
