using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

// Cash-flow functions with explicit dates (XNPV, XIRR). Values and dates are parallel ranges; the first
// date is the discounting anchor (dates may be out of order, but none may precede the first). XIRR uses the
// same robust bracketing solver as IRR/RATE and needs at least one sign change → #NUM! otherwise.

[MemoryPackable]
public sealed partial record XNpv(Expression[] Arguments) : Function
{
    // XNPV(rate, values, dates) — net present value of dated cash flows (365-day actual discounting).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (FinancialArguments.Number(Arguments, 0, context, out var rate) is { } rateError)
        {
            return ComputedValue.Error(rateError);
        }

        if (rate == -1)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (DatedFlows.Read(Arguments, 1, context, out var flows, out var dates) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return FinancialArguments.Result(BondMath.XNpv(rate, flows, dates));
    }
}

[MemoryPackable]
public sealed partial record XIrr(Expression[] Arguments) : Function
{
    // XIRR(values, dates, [guess]) — internal rate of return of dated cash flows. guess defaults to 0.1.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (DatedFlows.Read(Arguments, 0, context, out var flows, out var dates) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var guess = 0.1;
        if (Arguments.Length > 2 && FinancialArguments.Number(Arguments, 2, context, out guess) is { } guessError)
        {
            return ComputedValue.Error(guessError);
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
            return ComputedValue.Error(Error.Num);
        }

        return FinancialArguments.Result(BondMath.XIrr(flows, dates, guess));
    }
}

internal static class DatedFlows
{
    /// <summary>
    /// Reads the values (argument 0) and dates (argument 1) into aligned parallel lists. Returns an error if
    /// a cell fails to coerce, the two ranges differ in length, or a date precedes the anchor (first date).
    /// </summary>
    public static Error? Read(
        Expression[] arguments,
        int valuesIndex,
        EvaluationContext context,
        out List<double> flows,
        out List<DateTime> dates
    )
    {
        flows = new List<double>();
        dates = new List<DateTime>();

        var valueCells = ArgumentFlattening.ExpandComputedValues(arguments[valuesIndex], context);
        var dateCells = ArgumentFlattening.ExpandComputedValues(arguments[valuesIndex + 1], context);

        if (valueCells.Count != dateCells.Count)
        {
            return Error.Num;
        }

        for (var i = 0; i < valueCells.Count; i++)
        {
            if (valueCells[i].CoerceToNumber(out var flow) is { } flowError)
            {
                return flowError;
            }

            if (dateCells[i].CoerceToNumber(out var serial) is { } serialError)
            {
                return serialError;
            }

            if (DateSerial.ToDateTime(Math.Floor(serial), out var date) is { } rangeError)
            {
                return rangeError;
            }

            flows.Add(flow);
            dates.Add(date);
        }

        if (dates.Count == 0)
        {
            return Error.Num;
        }

        var anchor = dates[0];
        foreach (var date in dates)
        {
            if (date < anchor)
            {
                return Error.Num;
            }
        }

        return null;
    }
}
