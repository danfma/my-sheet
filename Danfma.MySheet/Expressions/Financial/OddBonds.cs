using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Odd first/last coupon bonds. These are the hardest formulas in the set; the arithmetic (BondMath) is a
// verbatim port of Excel's, validated against the oracle. ODDFPRICE/ODDFYIELD additionally require the odd
// first coupon to line up with the regular schedule stepping back from maturity → #NUM! otherwise.

[MemoryPackable]
public sealed partial record OddFPrice(Expression[] Arguments) : Function
{
    // ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis]).
    public override ComputedValue Evaluate(EvaluationContext context) => OddFirst.Evaluate(Arguments, context, yield: false);
}

[MemoryPackable]
public sealed partial record OddFYield(Expression[] Arguments) : Function
{
    // ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis]).
    public override ComputedValue Evaluate(EvaluationContext context) => OddFirst.Evaluate(Arguments, context, yield: true);
}

[MemoryPackable]
public sealed partial record OddLPrice(Expression[] Arguments) : Function
{
    // ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis]).
    public override ComputedValue Evaluate(EvaluationContext context) => OddLast.Evaluate(Arguments, context, yield: false);
}

[MemoryPackable]
public sealed partial record OddLYield(Expression[] Arguments) : Function
{
    // ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis]).
    public override ComputedValue Evaluate(EvaluationContext context) => OddLast.Evaluate(Arguments, context, yield: true);
}

internal static class OddFirst
{
    public static ComputedValue Evaluate(Expression[] arguments, EvaluationContext context, bool yield)
    {
        if (FinancialArguments.Date(arguments, 0, context, out var settlement) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(arguments, 1, context, out var maturity) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Date(arguments, 2, context, out var issue) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Date(arguments, 3, context, out var firstCoupon) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(arguments, 4, context, out var rate) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Number(arguments, 5, context, out var priceOrYield) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (FinancialArguments.Number(arguments, 6, context, out var redemption) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        if (FinancialArguments.Frequency(arguments, 7, context, out var frequency) is { } e7)
        {
            return ComputedValue.Error(e7);
        }

        if (FinancialArguments.Basis(arguments, 8, context, out var basis) is { } e8)
        {
            return ComputedValue.Error(e8);
        }

        if (
            !(issue < settlement && settlement < firstCoupon && firstCoupon < maturity)
            || rate < 0
            || priceOrYield < 0
            || redemption < 0
            || !BondMath.OddFirstCouponAligned(maturity, firstCoupon, frequency)
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        var value = yield
            ? BondMath.OddFYield(settlement, maturity, issue, firstCoupon, rate, priceOrYield, redemption, frequency, basis)
            : BondMath.OddFPrice(settlement, maturity, issue, firstCoupon, rate, priceOrYield, redemption, frequency, basis);

        return FinancialArguments.Result(value);
    }
}

internal static class OddLast
{
    public static ComputedValue Evaluate(Expression[] arguments, EvaluationContext context, bool yield)
    {
        if (FinancialArguments.Date(arguments, 0, context, out var settlement) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(arguments, 1, context, out var maturity) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Date(arguments, 2, context, out var lastInterest) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(arguments, 3, context, out var rate) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(arguments, 4, context, out var priceOrYield) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Number(arguments, 5, context, out var redemption) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (FinancialArguments.Frequency(arguments, 6, context, out var frequency) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        if (FinancialArguments.Basis(arguments, 7, context, out var basis) is { } e7)
        {
            return ComputedValue.Error(e7);
        }

        if (!(lastInterest < settlement && settlement < maturity) || rate < 0 || priceOrYield < 0 || redemption < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var value = BondMath.OddLFunc(
            settlement,
            maturity,
            lastInterest,
            rate,
            priceOrYield,
            redemption,
            frequency,
            basis,
            isLPrice: !yield
        );

        return FinancialArguments.Result(value);
    }
}
