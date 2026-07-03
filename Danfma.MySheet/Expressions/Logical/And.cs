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

        var result = true;

        foreach (var argument in Arguments)
        {
            if (argument.Evaluate(context).CoerceToBool(out var value) is { } error)
            {
                return ComputedValue.Error(error);
            }

            result &= value;
        }

        return ComputedValue.Boolean(result);
    }
}
