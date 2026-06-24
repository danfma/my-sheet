using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Max(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var (numbers, error) = NumericAggregation.Gather(Arguments, workbook);

        if (error is not null)
        {
            return error;
        }

        if (numbers.Count == 0)
        {
            return 0.0;
        }

        var result = numbers[0];

        foreach (var number in numbers)
        {
            if (number > result)
            {
                result = number;
            }
        }

        return result;
    }
}
