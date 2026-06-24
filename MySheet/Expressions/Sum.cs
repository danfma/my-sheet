using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Sum(Expression[] Expressions) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var fold = new SumFold();
        var error = NumericAggregation.Fold(Expressions, workbook, ref fold);

        return error ?? (object)fold.Total;
    }

    private struct SumFold : INumericFold
    {
        public double Total;

        public void Accept(double value) => Total += value;
    }
}
