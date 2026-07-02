using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

[MemoryPackable]
public sealed partial record TextJoin(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToText(out var delimiter) is { } delimiterError)
        {
            return ComputedValue.Error(delimiterError);
        }

        if (Arguments[1].Evaluate(context).CoerceToBool(out var ignoreEmpty) is { } ignoreError)
        {
            return ComputedValue.Error(ignoreError);
        }

        var parts = new List<string>();

        foreach (var value in ArgumentFlattening.FlattenComputedValues(Arguments[2..], context))
        {
            if (value.CoerceToText(out var text) is { } error)
            {
                return ComputedValue.Error(error);
            }

            if (ignoreEmpty && text.Length == 0)
            {
                continue;
            }

            parts.Add(text);
        }

        return ComputedValue.Text(string.Join(delimiter, parts));
    }
}
