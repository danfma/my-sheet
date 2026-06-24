using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record IfNa(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var value = Arguments[0].Compute(workbook);

        // Only #N/A is caught; other errors pass through.
        return value is ErrorValue { ErrorCode: "#N/A" } ? Arguments[1].Compute(workbook) : value;
    }
}
