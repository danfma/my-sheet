using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountA(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.Flatten(Arguments, context))
        {
            if (value is not null)
            {
                count++;
            }
        }

        return (double)count;
    }
}
