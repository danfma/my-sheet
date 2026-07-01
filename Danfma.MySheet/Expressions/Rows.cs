using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Rows(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Number(Arguments[0] is RangeReference range ? range.RowCount : 1.0);

}
