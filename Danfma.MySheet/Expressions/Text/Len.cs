using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

[MemoryPackable]
public sealed partial record Len(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(text.Length);
    }
}
