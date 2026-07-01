using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Lower(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Text(text.ToLowerInvariant());
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
