using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountA(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.Flatten(Arguments, workbook))
        {
            if (value is not null)
            {
                count++;
            }
        }

        return (double)count;
    }
}
