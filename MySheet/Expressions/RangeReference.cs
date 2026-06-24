using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record RangeReference : Reference
{
    public override object? Compute(Workbook workbook)
    {
        throw new NotImplementedException();
    }
}