using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Rate conversions, cumulative loan amounts and small value helpers. Arithmetic in BondMath.

[MemoryPackable]
public sealed partial record Effect(Expression[] Arguments) : Function
{
    // EFFECT(nominal_rate, npery) — effective annual interest rate.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var nominal) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var npery) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (nominal <= 0 || npery < 1)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Effect(nominal, npery));
    }
}

[MemoryPackable]
public sealed partial record Nominal(Expression[] Arguments) : Function
{
    // NOMINAL(effect_rate, npery) — nominal annual interest rate.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var effect) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var npery) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (effect <= 0 || npery < 1)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Nominal(effect, npery));
    }
}

[MemoryPackable]
public sealed partial record Mirr(Expression[] Arguments) : Function
{
    // MIRR(values, finance_rate, reinvest_rate) — modified internal rate of return.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet values range is a structural #REF!, before the flows are read.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var flows = new List<double>();
        if (FinancialArguments.NumberList(Arguments[0], context, flows) is { } flowError)
        {
            return ComputedValue.Error(flowError);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var financeRate) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var reinvestRate) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (flows.Count < 2)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var hasPositive = false;
        var hasNegative = false;
        foreach (var flow in flows)
        {
            hasPositive |= flow > 0;
            hasNegative |= flow < 0;
        }

        if (!hasPositive || !hasNegative)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return FinancialArguments.Result(BondMath.Mirr(flows, financeRate, reinvestRate));
    }
}

[MemoryPackable]
public sealed partial record Rri(Expression[] Arguments) : Function
{
    // RRI(nper, pv, fv) — equivalent interest rate for the growth of an investment.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var nper) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var pv) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var fv) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (nper <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (fv != pv && (pv == 0 || fv / pv < 0))
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.Rri(nper, pv, fv));
    }
}

[MemoryPackable]
public sealed partial record PDuration(Expression[] Arguments) : Function
{
    // PDURATION(rate, pv, fv) — periods required for an investment to reach a value.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var rate) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var pv) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var fv) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (rate <= 0 || pv <= 0 || fv <= 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.PDuration(rate, pv, fv));
    }
}

[MemoryPackable]
public sealed partial record ISPmt(Expression[] Arguments) : Function
{
    // ISPMT(rate, per, nper, pv) — interest paid during a specific period of a straight-loan.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var rate) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var per) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var nper) is { } e2)
        {
            return ComputedValue.Error(e2);
        }

        if (FinancialArguments.Number(Arguments, 3, context, out var pv) is { } e3)
        {
            return ComputedValue.Error(e3);
        }

        if (nper == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return FinancialArguments.Result(BondMath.ISPmt(rate, per, nper, pv));
    }
}

[MemoryPackable]
public sealed partial record CumIPmt(Expression[] Arguments) : Function
{
    // CUMIPMT(rate, nper, pv, start_period, end_period, type) — cumulative interest over a period range.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var rate) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var nper) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var pv) is { } e2)
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

        if (FinancialArguments.Number(Arguments, 5, context, out var type) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (
            rate <= 0
            || nper <= 0
            || pv <= 0
            || startPeriod < 1
            || endPeriod < startPeriod
            || endPeriod > nper
            || type is not (0 or 1)
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.CumIPmt(rate, nper, pv, startPeriod, endPeriod, (int)type));
    }
}

[MemoryPackable]
public sealed partial record CumPrinc(Expression[] Arguments) : Function
{
    // CUMPRINC(rate, nper, pv, start_period, end_period, type) — cumulative principal over a period range.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var rate) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var nper) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (FinancialArguments.Number(Arguments, 2, context, out var pv) is { } e2)
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

        if (FinancialArguments.Number(Arguments, 5, context, out var type) is { } e5)
        {
            return ComputedValue.Error(e5);
        }

        if (
            rate <= 0
            || nper <= 0
            || pv <= 0
            || startPeriod < 1
            || endPeriod < startPeriod
            || endPeriod > nper
            || type is not (0 or 1)
        )
        {
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.CumPrinc(rate, nper, pv, startPeriod, endPeriod, (int)type));
    }
}

[MemoryPackable]
public sealed partial record FvSchedule(Expression[] Arguments) : Function
{
    // FVSCHEDULE(principal, schedule) — future value after applying a series of compound rates.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet schedule range is a structural #REF!, before the rates are read.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (FinancialArguments.Number(Arguments, 0, context, out var principal) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        var schedule = new List<double>();
        if (FinancialArguments.NumberList(Arguments[1], context, schedule) is { } scheduleError)
        {
            return ComputedValue.Error(scheduleError);
        }

        return FinancialArguments.Result(BondMath.FvSchedule(principal, schedule));
    }
}

[MemoryPackable]
public sealed partial record DollarDe(Expression[] Arguments) : Function
{
    // DOLLARDE(fractional_dollar, fraction) — converts a fractional-notation price to a decimal.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var fractionalDollar) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var fraction) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (fraction < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (fraction is >= 0 and < 1)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return FinancialArguments.Result(BondMath.DollarDe(fractionalDollar, fraction));
    }
}

[MemoryPackable]
public sealed partial record DollarFr(Expression[] Arguments) : Function
{
    // DOLLARFR(decimal_dollar, fraction) — converts a decimal price to fractional notation.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var decimalDollar) is { } e0)
        {
            return ComputedValue.Error(e0);
        }

        if (FinancialArguments.Number(Arguments, 1, context, out var fraction) is { } e1)
        {
            return ComputedValue.Error(e1);
        }

        if (fraction < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (fraction is >= 0 and < 1)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        return FinancialArguments.Result(BondMath.DollarFr(decimalDollar, fraction));
    }
}
