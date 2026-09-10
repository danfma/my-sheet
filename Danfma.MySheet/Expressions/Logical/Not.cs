using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record Not(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // NOT's argument is a boolean condition slot, so it takes the same text "TRUE"/"FALSE" rule as IF's
        // condition. Registered Elementwise<Not>, so lifting over an array re-enters here per element.
        if (
            Arguments[0].Evaluate(context).CoerceToBoolAllowingTextWords(out var value) is { } error
        )
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Boolean(!value);
    }
}
