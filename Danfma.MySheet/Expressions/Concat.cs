using System.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Concat(Expression[] Arguments) : Function
{
    public override object? Compute(EvaluationContext context)
    {
        var builder = new StringBuilder();

        foreach (var value in ArgumentFlattening.Flatten(Arguments, context))
        {
            if (ValueCoercion.TryToText(value, out var text) is { } error)
            {
                return error;
            }

            builder.Append(text);
        }

        return builder.ToString();
    }
}
