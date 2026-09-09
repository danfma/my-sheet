using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Rows(Expression[] Arguments) : Function
{
    // A defined name that stands for a range counts its rows; a whole-column/row reference uses the
    // populated extent on its open row axis (structural on a bounded one); anything else (a single cell
    // or a scalar) is 1 — except a reference that FAILED to resolve, which reports its own error.
    // boundOpenRanges:false keeps the open reference so the extent rule applies.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF!, not an empty (0-row) extent.
        if (ReferenceGuard.MissingSheet(Arguments[0], context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (
            !NamedReferences.TryResolveReference(
                Arguments[0],
                context,
                out var reference,
                boundOpenRanges: false
            )
        )
        {
            // An argument that does not resolve reports its OWN error (#NAME? for an unknown name, #REF! for
            // a failed INDIRECT/OFFSET) instead of the plausible-looking 1 that used to hide a broken
            // reference; a scalar VALUE keeps that 1, since Excel counts it as a 1x1 array.
            return ReferencePosition.Unresolved(Arguments[0], context, ComputedValue.Number(1));
        }

        return ComputedValue.Number(
            reference switch
            {
                RangeReference range => range.RowCount,
                OpenRangeReference open => open.RowExtent(context),
                _ => 1.0,
            }
        );
    }
}
