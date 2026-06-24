using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Sum(Expression[] Expressions) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var (numbers, error) = NumericAggregation.Gather(Expressions, workbook);

        if (error is not null)
        {
            return error;
        }

        var total = 0.0;

        foreach (var number in numbers)
        {
            total += number;
        }

        return total;
    }
}
