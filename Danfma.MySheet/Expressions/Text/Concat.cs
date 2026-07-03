using System.Text;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Text;

[MemoryPackable]
public sealed partial record Concat(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet range is a structural #REF! — an open-range ghost would otherwise be read as empty
        // and concatenated into "".
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

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
