using MemoryPack;

namespace Danfma.MySheet.Expressions.Logical;

[MemoryPackable]
public sealed partial record IfError(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var value = ScalarReferenceValue.Evaluate(Arguments[0], context);

        // The fallback is only computed when the first argument is an error (short-circuit).
        return value.Kind == ComputedValueKind.Error ? Arguments[1].Evaluate(context) : value;
    }
}
