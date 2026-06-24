using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Count(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        // COUNT only tallies numeric values and, unlike SUM, never propagates errors.
        var (numbers, _) = NumericAggregation.Gather(Arguments, workbook);

        return (double)numbers.Count;
    }
}
