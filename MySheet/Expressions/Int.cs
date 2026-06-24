using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Int(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(workbook), out var number) is { } error)
        {
            return error;
        }

        return Math.Floor(number);
    }
}
