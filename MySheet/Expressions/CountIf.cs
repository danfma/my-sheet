using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountIf(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var criteria = Criteria.Parse(Arguments[1].Compute(workbook));
        var count = 0;

        foreach (var value in ArgumentFlattening.Expand(Arguments[0], workbook))
        {
            if (criteria.Matches(value))
            {
                count++;
            }
        }

        return (double)count;
    }
}
