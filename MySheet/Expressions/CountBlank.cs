using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountBlank(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.Flatten(Arguments, workbook))
        {
            if (value is null || value is string { Length: 0 })
            {
                count++;
            }
        }

        return (double)count;
    }
}
