using System.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Concat(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var builder = new StringBuilder();

        foreach (var value in ArgumentFlattening.Flatten(Arguments, context))
        {
            if (ValueCoercion.TryToText(value, out var text) is { } error)
            {
                return ComputedValue.From(error);
            }

            builder.Append(text);
        }

        return ComputedValue.Text(builder.ToString());
    }
}
