using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Rows(Expression[] Arguments) : Function
{
    // A defined name that stands for a range counts its rows; anything else (a single cell or a scalar) is 1.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Number(
            NamedReferences.TryResolveReference(Arguments[0], context, out var reference)
            && reference is RangeReference range
                ? range.RowCount
                : 1.0
        );
}
