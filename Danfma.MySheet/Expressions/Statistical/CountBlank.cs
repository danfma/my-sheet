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

        // The slot is opened through PositionalRange.Open so the criteria family's error-propagation arm
        // serves COUNTBLANK too (sweep item 34(a)): an error-valued argument — an unresolved name's
        // #NAME?, an unresolvable structured reference's #REF!, 1/0's #DIV/0! — propagates instead of
        // being scanned as the one non-blank element (a silent 0 on the oracle's error, measured
        // 2026-09-11, both entry modes). The scan itself is the element walk FlattenComputedValues
        // produced: cells for a reference, the node's own single value otherwise.
        foreach (var argument in Arguments)
        {
            if (PositionalRange.OpenCriteria(argument, context, out var range) is { } slotError)
            {
                return ComputedValue.Error(slotError);
            }

            for (var i = 0; i < range.Count; i++)
            {
                var value = range.Next();

                if (
                    value.Kind == ComputedValueKind.Blank
                    || (value.TryGetText(out var text) && text.Length == 0)
                )
                {
                    count++;
                }
            }
        }

        return ComputedValue.Number(count);
    }
}
