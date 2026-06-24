using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record IfError(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var value = Arguments[0].Compute(workbook);

        // The fallback is only computed when the first argument is an error (short-circuit).
        return value is ErrorValue ? Arguments[1].Compute(workbook) : value;
    }
}
