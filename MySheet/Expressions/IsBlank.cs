using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record IsBlank(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook) => Arguments[0].Compute(workbook) is null;
}
