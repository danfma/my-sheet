using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Trim(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var text) is { } error)
        {
            return ComputedValue.Error(error);
        }

        // Excel TRIM strips leading/trailing spaces and collapses internal runs of spaces to one.
        return ComputedValue.Text(string.Join(' ', text.Split(' ', StringSplitOptions.RemoveEmptyEntries)));
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
