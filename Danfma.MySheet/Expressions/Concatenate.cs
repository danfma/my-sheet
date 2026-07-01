using System.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Concatenate(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var builder = new StringBuilder();

        foreach (var value in ArgumentFlattening.FlattenComputedValues(Arguments, context))
        {
            if (value.CoerceToText(out var text) is { } error)
            {
                return ComputedValue.Error(error);
            }

            builder.Append(text);
        }

        return ComputedValue.Text(builder.ToString());
    }
}
