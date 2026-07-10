using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

// Dispersion and shape statistics. The variance folds live in StatisticsMath (shared with
// SUBTOTAL); the *A variants collect with NumericAggregation.FoldA (referenced text → 0,
// logicals → 1/0). The internal static Compute methods are reused by the Compatibility aliases
// (STDEV/STDEVP/VAR/VARP), which are distinct nodes so the un-parse keeps the legacy spelling.

[MemoryPackable]
public sealed partial record StDevS(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // STDEV.S(number1, …) — sample standard deviation ("n−1" method); fewer than 2 → #DIV/0!.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (StatisticsMath.Collect(arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.SampleVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(Math.Sqrt(variance));
    }
}

[MemoryPackable]
public sealed partial record StDevP(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // STDEV.P(number1, …) — population standard deviation ("n" method); no values → #DIV/0!.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (StatisticsMath.Collect(arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.PopulationVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(Math.Sqrt(variance));
    }
}

[MemoryPackable]
public sealed partial record StDevA(Expression[] Arguments) : Function
{
    // STDEVA(value1, …) — STDEV.S with the *A gathering rule.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.CollectA(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.SampleVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(Math.Sqrt(variance));
    }
}

[MemoryPackable]
public sealed partial record StDevPA(Expression[] Arguments) : Function
{
    // STDEVPA(value1, …) — STDEV.P with the *A gathering rule.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.CollectA(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.PopulationVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(Math.Sqrt(variance));
    }
}

[MemoryPackable]
public sealed partial record VarS(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // VAR.S(number1, …) — sample variance; fewer than 2 → #DIV/0!.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (StatisticsMath.Collect(arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.SampleVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(variance);
    }
}

[MemoryPackable]
public sealed partial record VarP(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // VAR.P(number1, …) — population variance; no values → #DIV/0!.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (StatisticsMath.Collect(arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.PopulationVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(variance);
    }
}

[MemoryPackable]
public sealed partial record VarA(Expression[] Arguments) : Function
{
    // VARA(value1, …) — VAR.S with the *A gathering rule.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.CollectA(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.SampleVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(variance);
    }
}

[MemoryPackable]
public sealed partial record VarPA(Expression[] Arguments) : Function
{
    // VARPA(value1, …) — VAR.P with the *A gathering rule.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.CollectA(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.PopulationVariance(values, out var variance) is { } varianceError
            ? ComputedValue.Error(varianceError)
            : ComputedValue.Number(variance);
    }
}

[MemoryPackable]
public sealed partial record AveDev(Expression[] Arguments) : Function
{
    // AVEDEV(number1, …) — mean of the absolute deviations from the mean; no values → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (values.Count == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var mean = StatisticsMath.Mean(values);
        var total = 0.0;

        foreach (var value in values)
        {
            total += Math.Abs(value - mean);
        }

        return ComputedValue.Number(total / values.Count);
    }
}

[MemoryPackable]
public sealed partial record DevSq(Expression[] Arguments) : Function
{
    // DEVSQ(number1, …) — sum of squared deviations from the mean; no values → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return values.Count == 0
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(StatisticsMath.SumSquaredDeviations(values));
    }
}

[MemoryPackable]
public sealed partial record GeoMean(Expression[] Arguments) : Function
{
    // GEOMEAN(number1, …) — geometric mean via exp(mean(ln x)); any value ≤ 0 (or no values) → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (values.Count == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var logTotal = 0.0;

        foreach (var value in values)
        {
            if (value <= 0)
            {
                return ComputedValue.Error(Error.Num);
            }

            logTotal += Math.Log(value);
        }

        return ComputedValue.Number(Math.Exp(logTotal / values.Count));
    }
}

[MemoryPackable]
public sealed partial record HarMean(Expression[] Arguments) : Function
{
    // HARMEAN(number1, …) — n / Σ(1/x); any value ≤ 0 (or no values) → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (values.Count == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var reciprocalTotal = 0.0;

        foreach (var value in values)
        {
            if (value <= 0)
            {
                return ComputedValue.Error(Error.Num);
            }

            reciprocalTotal += 1 / value;
        }

        return ComputedValue.Number(values.Count / reciprocalTotal);
    }
}

[MemoryPackable]
public sealed partial record Skew(Expression[] Arguments) : Function
{
    // SKEW(number1, …) — sample skewness: n/((n−1)(n−2)) · Σ((x−x̄)/s)³ with s the sample standard
    // deviation; fewer than 3 points or s = 0 → #DIV/0!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var n = values.Count;

        if (
            n < 3
            || StatisticsMath.SampleVariance(values, out var variance) is not null
            || variance == 0
        )
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var mean = StatisticsMath.Mean(values);
        var s = Math.Sqrt(variance);
        var total = 0.0;

        foreach (var value in values)
        {
            var z = (value - mean) / s;
            total += z * z * z;
        }

        return ComputedValue.Number((double)n / ((n - 1) * (n - 2)) * total);
    }
}

[MemoryPackable]
public sealed partial record SkewP(Expression[] Arguments) : Function
{
    // SKEW.P(number1, …) — population skewness: (1/n) · Σ((x−μ)/σ)³ with the population standard
    // deviation; fewer than 3 points or σ = 0 → #DIV/0!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var n = values.Count;

        if (
            n < 3
            || StatisticsMath.PopulationVariance(values, out var variance) is not null
            || variance == 0
        )
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var mean = StatisticsMath.Mean(values);
        var sigma = Math.Sqrt(variance);
        var total = 0.0;

        foreach (var value in values)
        {
            var z = (value - mean) / sigma;
            total += z * z * z;
        }

        return ComputedValue.Number(total / n);
    }
}

[MemoryPackable]
public sealed partial record Kurt(Expression[] Arguments) : Function
{
    // KURT(number1, …) — excess kurtosis, Excel's sample formula:
    // n(n+1)/((n−1)(n−2)(n−3)) · Σ((x−x̄)/s)⁴ − 3(n−1)²/((n−2)(n−3));
    // fewer than 4 points or s = 0 → #DIV/0!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (StatisticsMath.Collect(Arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var n = values.Count;

        if (
            n < 4
            || StatisticsMath.SampleVariance(values, out var variance) is not null
            || variance == 0
        )
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var mean = StatisticsMath.Mean(values);
        var s = Math.Sqrt(variance);
        var total = 0.0;

        foreach (var value in values)
        {
            var z = (value - mean) / s;
            total += z * z * z * z;
        }

        var lead = (double)n * (n + 1) / ((n - 1) * (n - 2) * (n - 3));
        var tail = 3.0 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));

        return ComputedValue.Number(lead * total - tail);
    }
}

[MemoryPackable]
public sealed partial record Standardize(Expression[] Arguments) : Function
{
    // STANDARDIZE(x, mean, standard_dev) — the z-score (x − mean)/standard_dev; sd ≤ 0 → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var x) is { } xError)
        {
            return ComputedValue.Error(xError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var mean) is { } meanError)
        {
            return ComputedValue.Error(meanError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var sd) is { } sdError)
        {
            return ComputedValue.Error(sdError);
        }

        return sd <= 0 ? ComputedValue.Error(Error.Num) : ComputedValue.Number((x - mean) / sd);
    }
}
