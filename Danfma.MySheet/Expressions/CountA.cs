using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.Flatten(Arguments, context))
        {
            if (value is not null)
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
