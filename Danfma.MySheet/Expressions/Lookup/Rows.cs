using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Rows(Expression[] Arguments) : Function
{
    // A defined name that stands for a range counts its rows; a whole-column/row reference uses the
    // populated extent on its open row axis (structural on a bounded one); anything else (a single cell
    // or a scalar) is 1 — except a reference that FAILED to resolve, which reports its own error.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF!, not an empty (0-row) extent. This SYNTACTIC
        // pass is the cheap short-circuit for a ghost written on the argument node itself (and the only pass
        // that answers #REF! for a ':' range whose endpoints cannot form one at all); the resolved target's
        // own sheet is re-checked inside ReferencePosition.TryResolve below, which is what catches a ghost
        // reached THROUGH a function — ROWS(INDEX(Ghost!A1:A3,2,1)).
        if (ReferenceGuard.MissingSheet(Arguments[0], context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // A COMPUTED array — a producer (FILTER/SORT/UNIQUE/SEQUENCE), an operator over a range, a lifted
        // function, IF — answers its own row count through the shared consumer gate, which keeps a bare
        // reference or name on the reference path below. This must come AFTER the syntactic guard above
        // (a ghost source is structural #REF!, not a shape) and BEFORE TryResolve, whose failure arm would
        // otherwise re-evaluate the array to its collapsed top-left and report 1 or that element's error.
        if (ArrayEvaluation.TryStream(Arguments[0], context, out var array))
        {
            return ReferencePosition.ArrayExtent(array, array.Rows);
        }

        // The fallback is 1 (Excel counts a scalar as a 1x1 array); an argument that does not resolve reports
        // its OWN error instead — see ReferencePosition.TryResolve for both failure arms.
        if (
            !ReferencePosition.TryResolve(
                Arguments[0],
                context,
                ComputedValue.Number(1),
                out var reference,
                out var failure
            )
        )
        {
            return failure;
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
