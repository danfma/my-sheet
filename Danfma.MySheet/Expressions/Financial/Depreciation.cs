using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Depreciation functions. The arithmetic lives in BondMath (a faithful, oracle-validated port of Excel's
// conventions); these records handle argument coercion, optional defaults and Excel's domain errors.

[MemoryPackable]
public sealed partial record Sln(Expression[] Arguments) : Function
{
    // SLN(cost, salvage, life) — straight-line depreciation per period.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var salvage) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var life) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (life == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return FinancialArguments.Result(BondMath.Sln(cost, salvage, life));
    }
}

[MemoryPackable]
public sealed partial record Syd(Expression[] Arguments) : Function
{
    // SYD(cost, salvage, life, per) — sum-of-years'-digits depreciation for period `per`.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var salvage) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var life) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var per) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (life <= 0 || per < 1 || per > life)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Syd(cost, salvage, life, per));
    }
}

[MemoryPackable]
public sealed partial record Db(Expression[] Arguments) : Function
{
    // DB(cost, salvage, life, period, [month]) — fixed-declining-balance depreciation. month defaults to 12.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var salvage) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var life) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var period) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        var month = 12.0;
        if (Arguments.Length > 4 && FinancialArguments.Number(Arguments, 4, context, out month) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (cost < 0 || salvage < 0 || life <= 0 || month is < 1 or > 12 || period < 1 || period > life)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Db(cost, salvage, life, period, month));
    }
}

[MemoryPackable]
public sealed partial record Ddb(Expression[] Arguments) : Function
{
    // DDB(cost, salvage, life, period, [factor]) — double-declining-balance depreciation. factor defaults to 2.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var salvage) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var life) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var period) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        var factor = 2.0;
        if (Arguments.Length > 4 && FinancialArguments.Number(Arguments, 4, context, out factor) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (cost < 0 || salvage < 0 || life <= 0 || factor <= 0 || period < 1 || period > life)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Ddb(cost, salvage, life, period, factor));
    }
}

[MemoryPackable]
public sealed partial record Vdb(Expression[] Arguments) : Function
{
    // VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch]) — variable declining balance.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var salvage) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var life) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var startPeriod) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(Arguments, 4, context, out var endPeriod) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        var factor = 2.0;
        if (Arguments.Length > 5 && FinancialArguments.Number(Arguments, 5, context, out factor) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        var noSwitch = false;
        if (Arguments.Length > 6 && Arguments[6].Evaluate(context).CoerceToBool(out noSwitch) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        if (
            cost < 0
            || salvage < 0
            || life <= 0
            || factor <= 0
            || startPeriod < 0
            || endPeriod <= 0
            || startPeriod > endPeriod
            || startPeriod > life
            || endPeriod > life
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Vdb(cost, salvage, life, startPeriod, endPeriod, factor, noSwitch));
    }
}

[MemoryPackable]
public sealed partial record AmorLinc(Expression[] Arguments) : Function
{
    // AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis]) — French linear
    // depreciation. basis 2 (actual/360) is rejected by Excel → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(Arguments, 1, context, out var datePurchased) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Date(Arguments, 2, context, out var firstPeriod) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var salvage) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(Arguments, 4, context, out var period) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Number(Arguments, 5, context, out var rate) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (FinancialArguments.Basis(Arguments, 6, context, out var basis) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        if (
            cost < 0
            || salvage < 0
            || salvage >= cost
            || period < 0
            || rate < 0
            || datePurchased >= firstPeriod
            || basis == 2
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(
            BondMath.AmorLinc(cost, datePurchased, firstPeriod, salvage, period, rate, basis)
        );
    }
}

[MemoryPackable]
public sealed partial record AmorDegrc(Expression[] Arguments) : Function
{
    // AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis]) — French declining
    // depreciation with a life-dependent coefficient. Excel rounds each step; basis 2 is rejected → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var cost) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Date(Arguments, 1, context, out var datePurchased) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Date(Arguments, 2, context, out var firstPeriod) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var salvage) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (FinancialArguments.Number(Arguments, 4, context, out var period) is { } e4)
        {
            return ComputedValue.Error(e4);
        }

        if (FinancialArguments.Number(Arguments, 5, context, out var rate) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (FinancialArguments.Basis(Arguments, 6, context, out var basis) is { } e6)
        {
            return ComputedValue.Error(e6);
        }

        // Excel rejects an implied asset life (1/rate) in (0,3] or (4,5) and basis 2 (actual/360).
        var assetLife = 1.0 / rate;
        if (
            rate <= 0
            || cost < 0
            || salvage < 0
            || salvage >= cost
            || period < 0
            || datePurchased >= firstPeriod
            || basis == 2
            || (assetLife > 0 && assetLife <= 3)
            || (assetLife >= 4 && assetLife <= 5)
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(
            BondMath.AmorDegrc(cost, datePurchased, firstPeriod, salvage, period, rate, basis, excelCompliant: true)
        );
    }
}
