using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountBlank(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF!, not an empty range of blanks.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var count = 0;

        foreach (var value in ArgumentFlattening.FlattenComputedValues(Arguments, context))
        {
            if (
                value.Kind == ComputedValueKind.Blank
                || (value.TryGetText(out var text) && text.Length == 0)
            )
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
