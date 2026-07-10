using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Periodic-coupon bond analytics: PRICE, YIELD, DURATION, MDURATION. YIELD inverts PRICE with the shared
// bracketing solver; the schedule and day counts come from BondMath.

[MemoryPackable]
public sealed partial record Price(Expression[] Arguments) : Function
{
    // PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis]) — price per $100 face value.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Date(Arguments, 0, context, out var settlement) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(Arguments, 1, context, out var maturity) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var rate) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var yld) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(Arguments, 4, context, out var redemption) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Frequency(Arguments, 5, context, out var frequency) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (FinancialArguments.Basis(Arguments, 6, context, out var basis) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        if (settlement >= maturity || rate < 0 || yld < 0 || redemption <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(
            BondMath.Price(settlement, maturity, rate, yld, redemption, frequency, basis)
        );
    }
}

[MemoryPackable]
public sealed partial record Yield(Expression[] Arguments) : Function
{
    // YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis]) — yield to maturity.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Date(Arguments, 0, context, out var settlement) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(Arguments, 1, context, out var maturity) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var rate) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var price) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(Arguments, 4, context, out var redemption) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Frequency(Arguments, 5, context, out var frequency) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (FinancialArguments.Basis(Arguments, 6, context, out var basis) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        if (settlement >= maturity || rate < 0 || price <= 0 || redemption <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(
            BondMath.Yield(settlement, maturity, rate, price, redemption, frequency, basis)
        );
    }
}

[MemoryPackable]
public sealed partial record Duration(Expression[] Arguments) : Function
{
    // DURATION(settlement, maturity, coupon, yld, frequency, [basis]) — Macaulay duration in years.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        DurationMath.Evaluate(Arguments, context, modified: false);
}

[MemoryPackable]
public sealed partial record MDuration(Expression[] Arguments) : Function
{
    // MDURATION(settlement, maturity, coupon, yld, frequency, [basis]) — modified Macaulay duration.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        DurationMath.Evaluate(Arguments, context, modified: true);
}

internal static class DurationMath
{
    public static ComputedValue Evaluate(
        Expression[] arguments,
        EvaluationContext context,
        bool modified
    )
    {
        if (FinancialArguments.Date(arguments, 0, context, out var settlement) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(arguments, 1, context, out var maturity) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(arguments, 2, context, out var coupon) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(arguments, 3, context, out var yld) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Frequency(arguments, 4, context, out var frequency) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Basis(arguments, 5, context, out var basis) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (settlement >= maturity || coupon < 0 || yld < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(
            BondMath.Duration(settlement, maturity, coupon, yld, frequency, basis, modified)
        );
    }
}
