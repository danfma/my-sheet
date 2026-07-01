using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountBlank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var count = 0;

        foreach (var value in ArgumentFlattening.FlattenComputedValues(Arguments, context))
        {
            if (value.Kind == ComputedValueKind.Blank || (value.TryGetText(out var text) && text.Length == 0))
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
