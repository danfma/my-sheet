using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Or(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var result = false;

        foreach (var argument in Arguments)
        {
            if (ValueCoercion.TryToBool(argument.Compute(workbook), out var value) is { } error)
            {
                return error;
            }

            result |= value;
        }

        return result;
    }
}
