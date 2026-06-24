using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record SumIf(Expression[] Arguments) : Function
{
    // SUMIF(range, criteria, [sum_range]) — sums sum_range (or range) where range matches the criteria.
    public override object? Compute(Workbook workbook)
    {
        var criteria = Criteria.Parse(Arguments[1].Compute(workbook));
        var range = ArgumentFlattening.Expand(Arguments[0], workbook);
        var sumRange = Arguments.Length == 3 ? ArgumentFlattening.Expand(Arguments[2], workbook) : range;

        var total = 0.0;
        var length = Math.Min(range.Count, sumRange.Count);

        for (var i = 0; i < length; i++)
        {
            if (criteria.Matches(range[i]) && sumRange[i] is double number)
            {
                total += number;
            }
        }

        return total;
    }
}
