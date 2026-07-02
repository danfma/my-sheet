using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountA(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.FlattenComputedValues(Arguments, context))
        {
            if (value.Kind != ComputedValueKind.Blank)
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
