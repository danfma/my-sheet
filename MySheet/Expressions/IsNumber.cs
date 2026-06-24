using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record IsNumber(Expression[] Arguments) : Function
{
    // An error value is "not a number" rather than propagated, matching Excel's IS functions.
    public override object? Compute(Workbook workbook) => Arguments[0].Compute(workbook) is double;
}
