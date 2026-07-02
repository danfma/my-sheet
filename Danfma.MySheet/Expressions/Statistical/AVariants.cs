using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

// The *A variants of the basic aggregates: same shape as AVERAGE/MAX/MIN but referenced text
// counts as 0 and referenced logicals as 1/0 (NumericAggregation.FoldA). Blanks stay ignored.

[MemoryPackable]
public sealed partial record AverageA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.CollectA(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return values.Count == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(StatisticsMath.Mean(values));
    }
}

[MemoryPackable]
public sealed partial record MaxA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new ExtremeFold { Larger = true };

        return NumericAggregation.FoldA(Arguments, context, ref fold) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(fold.HasValue ? fold.Value : 0.0);
    }
}

[MemoryPackable]
public sealed partial record MinA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var fold = new ExtremeFold { Larger = false };

        return NumericAggregation.FoldA(Arguments, context, ref fold) is { } error
            ? ComputedValue.Error(error)
            : ComputedValue.Number(fold.HasValue ? fold.Value : 0.0);
    }
}

file struct ExtremeFold : INumericFold
{
    public bool Larger;
    public bool HasValue;
    public double Value;

    public void Accept(double value)
    {
        if (!HasValue || (Larger ? value > Value : value < Value))
        {
            Value = value;
            HasValue = true;
        }
    }
}
