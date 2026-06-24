using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record TextJoin(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToText(Arguments[0].Compute(context), out var delimiter) is { } delimiterError)
        {
            return delimiterError;
        }

        if (ValueCoercion.TryToBool(Arguments[1].Compute(context), out var ignoreEmpty) is { } ignoreError)
        {
            return ignoreError;
        }

        var parts = new List<string>();

        foreach (var value in ArgumentFlattening.Flatten(Arguments[2..], context))
        {
            if (ValueCoercion.TryToText(value, out var text) is { } error)
            {
                return error;
            }

            if (ignoreEmpty && text.Length == 0)
            {
                continue;
            }

            parts.Add(text);
        }

        return string.Join(delimiter, parts);
    }
}
