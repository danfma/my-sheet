using MemoryPack;

namespace Danfma.MySheet.Expressions.Dates;

// Volatile clock functions (F1): NOW and TODAY. Both read the workbook's injectable TimeProvider in LOCAL
// time (like Excel) and return an Excel serial (OADate). They mark the evaluation volatile so the cell — and
// its dependents, transitively — is cached per epoch and refreshed by Workbook.Recalculate(); within an epoch
// the clock is sampled once (Workbook.EpochNow), so every NOW()/TODAY() in a pass agrees. Zero arguments.

[MemoryPackable]
public sealed partial record Now(Expression[] Arguments) : Function
{
    public override bool IsVolatile => true;

    public override ComputedValue Evaluate(EvaluationContext context)
    {
        context.Workbook.MarkVolatileTouched();
        return ComputedValue.Number(context.Workbook.EpochNow());
    }
}

[MemoryPackable]
public sealed partial record Today(Expression[] Arguments) : Function
{
    public override bool IsVolatile => true;

    public override ComputedValue Evaluate(EvaluationContext context)
    {
        context.Workbook.MarkVolatileTouched();
        // The whole-day part of the epoch's local NOW serial.
        return ComputedValue.Number(Math.Floor(context.Workbook.EpochNow()));
    }
}
