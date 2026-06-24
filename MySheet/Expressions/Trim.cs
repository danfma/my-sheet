using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Trim(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(context), out var text) is { } error)
        {
            return error;
        }

        // Excel TRIM strips leading/trailing spaces and collapses internal runs of spaces to one.
        return string.Join(' ', text.Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }
}
