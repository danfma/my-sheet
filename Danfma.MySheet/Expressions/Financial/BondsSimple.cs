using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Accrued interest, single-period discount/interest securities and T-bills. Closed-form math in BondMath.

[MemoryPackable]
public sealed partial record AccrInt(Expression[] Arguments) : Function
{
    // ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]).
    // calc_method defaults to TRUE (accrue from issue to settlement).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Date(Arguments, 0, context, out var issue) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(Arguments, 1, context, out var firstInterest) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Date(Arguments, 2, context, out var settlement) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var rate) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(Arguments, 4, context, out var par) is { } e4)
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

        var calcMethod = true;
        if (
            Arguments.Length > 7
            && Arguments[7].Evaluate(context).CoerceToBool(out calcMethod) is { } e7
        )
        {
            return ComputedValue.Error(e7);
        }

        if (rate <= 0 || par <= 0 || settlement <= issue || firstInterest < settlement)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(
            BondMath.AccrInt(
                issue,
                firstInterest,
                settlement,
                rate,
                par,
                frequency,
                basis,
                calcMethod ? 1 : 0
            )
        );
    }
}

[MemoryPackable]
public sealed partial record AccrIntM(Expression[] Arguments) : Function
{
    // ACCRINTM(issue, settlement, rate, par, [basis]) — accrued interest for a maturity-paying security.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Date(Arguments, 0, context, out var issue) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(Arguments, 1, context, out var settlement) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var rate) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var par) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Basis(Arguments, 4, context, out var basis) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (rate <= 0 || par <= 0 || settlement <= issue)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.AccrIntM(issue, settlement, rate, par, basis));
    }
}

[MemoryPackable]
public sealed partial record Disc(Expression[] Arguments) : Function
{
    // DISC(settlement, maturity, pr, redemption, [basis]) — discount rate of a security.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            SimpleBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var v1,
                out var v2,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (v1 <= 0 || v2 <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Disc(s, m, v1, v2, basis));
    }
}

[MemoryPackable]
public sealed partial record IntRate(Expression[] Arguments) : Function
{
    // INTRATE(settlement, maturity, investment, redemption, [basis]) — interest rate of a discount security.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            SimpleBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var investment,
                out var redemption,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (investment <= 0 || redemption <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.IntRate(s, m, investment, redemption, basis));
    }
}

[MemoryPackable]
public sealed partial record Received(Expression[] Arguments) : Function
{
    // RECEIVED(settlement, maturity, investment, discount, [basis]) — amount received at maturity.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            SimpleBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var investment,
                out var discount,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (investment <= 0 || discount <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Received(s, m, investment, discount, basis));
    }
}

[MemoryPackable]
public sealed partial record PriceDisc(Expression[] Arguments) : Function
{
    // PRICEDISC(settlement, maturity, discount, redemption, [basis]) — price of a discounted security.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            SimpleBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var discount,
                out var redemption,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (discount <= 0 || redemption <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.PriceDisc(s, m, discount, redemption, basis));
    }
}

[MemoryPackable]
public sealed partial record YieldDisc(Expression[] Arguments) : Function
{
    // YIELDDISC(settlement, maturity, pr, redemption, [basis]) — annual yield of a discounted security.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            SimpleBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var pr,
                out var redemption,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (pr <= 0 || redemption <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.YieldDisc(s, m, pr, redemption, basis));
    }
}

[MemoryPackable]
public sealed partial record PriceMat(Expression[] Arguments) : Function
{
    // PRICEMAT(settlement, maturity, issue, rate, yld, [basis]) — price of an interest-at-maturity security.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            MaturityBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var issue,
                out var rate,
                out var v,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (rate <= 0 || v <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.PriceMat(s, m, issue, rate, v, basis));
    }
}

[MemoryPackable]
public sealed partial record YieldMat(Expression[] Arguments) : Function
{
    // YIELDMAT(settlement, maturity, issue, rate, pr, [basis]) — annual yield of an interest-at-maturity bond.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (
            MaturityBond.Read(
                Arguments,
                context,
                out var s,
                out var m,
                out var issue,
                out var rate,
                out var pr,
                out var basis
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        if (rate <= 0 || pr <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.YieldMat(s, m, issue, rate, pr, basis));
    }
}

[MemoryPackable]
public sealed partial record TBillEq(Expression[] Arguments) : Function
{
    // TBILLEQ(settlement, maturity, discount) — bond-equivalent yield of a Treasury bill.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TBill.Evaluate(Arguments, context, TBillKind.Eq);
}

[MemoryPackable]
public sealed partial record TBillPrice(Expression[] Arguments) : Function
{
    // TBILLPRICE(settlement, maturity, discount) — price per $100 face value of a Treasury bill.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TBill.Evaluate(Arguments, context, TBillKind.Price);
}

[MemoryPackable]
public sealed partial record TBillYield(Expression[] Arguments) : Function
{
    // TBILLYIELD(settlement, maturity, pr) — yield of a Treasury bill.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        TBill.Evaluate(Arguments, context, TBillKind.Yield);
}

internal enum TBillKind
{
    Eq,
    Price,
    Yield,
}

internal static class TBill
{
    public static ComputedValue Evaluate(
        Expression[] arguments,
        EvaluationContext context,
        TBillKind kind
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

        if (FinancialArguments.Number(arguments, 2, context, out var third) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        // Excel caps a T-bill at one year from settlement.
        if (settlement >= maturity || maturity > settlement.AddYears(1) || third <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var value = kind switch
        {
            TBillKind.Eq => BondMath.TBillEq(settlement, maturity, third),
            TBillKind.Price => BondMath.TBillPrice(settlement, maturity, third),
            _ => BondMath.TBillYield(settlement, maturity, third),
        };

        return FinancialArguments.Result(value);
    }
}

// Shared readers for the (settlement, maturity, x, y, [basis]) and (settlement, maturity, issue, rate, y,
// [basis]) shapes, validating settlement < maturity (and issue ordering) up front.
internal static class SimpleBond
{
    public static Error? Read(
        Expression[] arguments,
        EvaluationContext context,
        out DateTime settlement,
        out DateTime maturity,
        out double value1,
        out double value2,
        out int basis
    )
    {
        maturity = default;
        value1 = 0;
        value2 = 0;
        basis = 0;

        if (FinancialArguments.Date(arguments, 0, context, out settlement) is { } e0)
        {
            return e0;
        }

        if (FinancialArguments.Date(arguments, 1, context, out maturity) is { } e1)
        {
            return e1;
        }

        if (FinancialArguments.Number(arguments, 2, context, out value1) is { } e2)
        {
            return e2;
        }

        if (FinancialArguments.Number(arguments, 3, context, out value2) is { } e3)
        {
            return e3;
        }

        if (FinancialArguments.Basis(arguments, 4, context, out basis) is { } e4)
        {
            return e4;
        }

        return settlement >= maturity ? Error.Num : null;
    }
}

internal static class MaturityBond
{
    public static Error? Read(
        Expression[] arguments,
        EvaluationContext context,
        out DateTime settlement,
        out DateTime maturity,
        out DateTime issue,
        out double rate,
        out double value,
        out int basis
    )
    {
        maturity = default;
        issue = default;
        rate = 0;
        value = 0;
        basis = 0;

        if (FinancialArguments.Date(arguments, 0, context, out settlement) is { } e0)
        {
            return e0;
        }

        if (FinancialArguments.Date(arguments, 1, context, out maturity) is { } e1)
        {
            return e1;
        }

        if (FinancialArguments.Date(arguments, 2, context, out issue) is { } e2)
        {
            return e2;
        }

        if (FinancialArguments.Number(arguments, 3, context, out rate) is { } e3)
        {
            return e3;
        }

        if (FinancialArguments.Number(arguments, 4, context, out value) is { } e4)
        {
            return e4;
        }

        if (FinancialArguments.Basis(arguments, 5, context, out basis) is { } e5)
        {
            return e5;
        }

        return settlement >= maturity || settlement <= issue || maturity <= issue
            ? Error.Num
            : null;
    }
}
