using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Average(Expression[] Arguments) : Function
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
            return ErrorValue.DivByZero;
        }

        var total = 0.0;

        foreach (var number in numbers)
        {
            total += number;
        }

        return total / numbers.Count;
    }
}
