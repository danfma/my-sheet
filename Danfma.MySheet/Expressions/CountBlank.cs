using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountBlank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.Flatten(Arguments, context))
        {
            if (value is null || value is string { Length: 0 })
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
