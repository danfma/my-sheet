using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

// Volatile RNG functions (F1): RAND and RANDBETWEEN. Both draw from the workbook's persistent, seedable RNG
// (Workbook.NextRandom), which also marks the evaluation volatile so the cell — and its dependents,
// transitively — is cached per epoch and refreshed by Workbook.Recalculate(). RAND() is a real number in
// [0, 1) (official RAND page). RANDBETWEEN(bottom, top) is an inclusive integer draw on both ends; the page
// is silent on reversed bounds and non-integer inputs, so (per the plan) bottom>top => #NUM! and non-integer
// bounds truncate toward zero.

[MemoryPackable]
public sealed partial record Rand(Expression[] Arguments) : Function
{
    public override bool IsVolatile => true;

    public override ComputedValue Evaluate(EvaluationContext context) =>
        ComputedValue.Number(context.Workbook.NextRandom());
}

[MemoryPackable]
public sealed partial record RandBetween(Expression[] Arguments) : Function
{
    public override bool IsVolatile => true;

    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var bottomArg) is { } bottomError)
        {
            return ComputedValue.Error(bottomError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var topArg) is { } topError)
        {
            return ComputedValue.Error(topError);
        }

        var bottom = Math.Truncate(bottomArg);
        var top = Math.Truncate(topArg);

        if (bottom > top)
        {
            return ComputedValue.Error(Error.Num);
        }

        // NextRandom() is [0, 1); scale to the inclusive integer range [bottom, top].
        var span = top - bottom + 1d;
        var draw = bottom + Math.Floor(context.Workbook.NextRandom() * span);

        // Defensive clamp against a floating-point rounding that lands one past the top.
        return ComputedValue.Number(draw > top ? top : draw);
    }
}
