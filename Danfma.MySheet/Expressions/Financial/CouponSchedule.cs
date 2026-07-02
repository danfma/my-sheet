using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Coupon-schedule functions: COUPPCD/COUPNCD (previous/next coupon dates, walking backward from maturity),
// COUPNUM (coupons until maturity) and the three day counts COUPDAYBS/COUPDAYS/COUPDAYSNC. All share the
// (settlement, maturity, frequency, [basis]) prefix and the settlement < maturity domain rule.

[MemoryPackable]
public sealed partial record CoupPcd(Expression[] Arguments) : Function
{
    // COUPPCD(settlement, maturity, frequency, [basis]) — previous coupon date before settlement (serial).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.CouponInputs(Arguments, context, out var s, out var m, out var f, out _) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(DateSerial.FromDateTime(BondMath.CoupPcd(s, m, f)));
    }
}

[MemoryPackable]
public sealed partial record CoupNcd(Expression[] Arguments) : Function
{
    // COUPNCD(settlement, maturity, frequency, [basis]) — next coupon date after settlement (serial).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.CouponInputs(Arguments, context, out var s, out var m, out var f, out _) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(DateSerial.FromDateTime(BondMath.CoupNcd(s, m, f)));
    }
}

[MemoryPackable]
public sealed partial record CoupNum(Expression[] Arguments) : Function
{
    // COUPNUM(settlement, maturity, frequency, [basis]) — number of coupons between settlement and maturity.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.CouponInputs(Arguments, context, out var s, out var m, out var f, out _) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return FinancialArguments.Result(BondMath.CoupNum(s, m, f));
    }
}

[MemoryPackable]
public sealed partial record CoupDays(Expression[] Arguments) : Function
{
    // COUPDAYS(settlement, maturity, frequency, [basis]) — days in the coupon period containing settlement.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.CouponInputs(Arguments, context, out var s, out var m, out var f, out var b) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return FinancialArguments.Result(BondMath.CoupDays(b, s, m, f));
    }
}

[MemoryPackable]
public sealed partial record CoupDayBs(Expression[] Arguments) : Function
{
    // COUPDAYBS(settlement, maturity, frequency, [basis]) — days from the period start to settlement.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.CouponInputs(Arguments, context, out var s, out var m, out var f, out var b) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return FinancialArguments.Result(BondMath.CoupDaysBs(b, s, m, f));
    }
}

[MemoryPackable]
public sealed partial record CoupDaysNc(Expression[] Arguments) : Function
{
    // COUPDAYSNC(settlement, maturity, frequency, [basis]) — days from settlement to the next coupon.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.CouponInputs(Arguments, context, out var s, out var m, out var f, out var b) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return FinancialArguments.Result(BondMath.CoupDaysNc(b, s, m, f));
    }
}
