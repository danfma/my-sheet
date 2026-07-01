using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Value(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // Reuse numeric coercion: numeric text parses, non-numeric text → #VALUE!.
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(number);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
