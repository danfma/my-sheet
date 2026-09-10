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

        // COUNTBLANK is defined over a RANGE, and the oracle rejects a computed array in its slot the way the
        // criteria family does (Phase 11a's Rule B): #REF! for a producer in both entry modes, #REF!
        // array-entered for a bare array expression (Aspose.Cells 26.6.0, 2026-09-10; DynamicArrayTests).
        // So the shared walk's streaming arm must never be reached here — a streamed blank would be counted.
        foreach (var argument in Arguments)
        {
            if (PositionalRange.RejectComputedArray(argument, context) is { } rejected)
            {
                return ComputedValue.Error(rejected);
            }
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
