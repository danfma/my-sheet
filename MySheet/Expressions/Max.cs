using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Max(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        var fold = new MaxFold();
        var error = NumericAggregation.Fold(Arguments, workbook, ref fold);

        return error ?? (object)(fold.HasValue ? fold.Value : 0.0);
    }

    private struct MaxFold : INumericFold
    {
        public bool HasValue;
        public double Value;

        public void Accept(double value)
        {
            if (!HasValue || value > Value)
            {
                Value = value;
                HasValue = true;
            }
        }
    }
}
