using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record And(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference argument to a missing sheet is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // Excel semantics (shared with OR/XOR): text and blank cells reached through a reference/array
        // are ignored; #VALUE! only when nothing evaluable survives. AND is TRUE iff every evaluable
        // logical value is TRUE.
        if (LogicalReduction.Reduce(Arguments, context, out var trueCount, out var total) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return total == 0 ? ComputedValue.Error(Error.Value) : ComputedValue.Boolean(trueCount == total);
    }
}
