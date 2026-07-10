using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

// The two-range statistics, all built on PairwiseRanges: positions where EITHER side is
// non-numeric are dropped as a pair (the documented behaviour), length mismatches → #N/A, cell
// errors propagate. The moment computations are shared through BivariateMoments; the internal
// static Compute methods are reused by the Compatibility aliases (COVAR, FORECAST).

[MemoryPackable]
public sealed partial record Correl(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        BivariateMoments.Correlation(Arguments, context);
}

[MemoryPackable]
public sealed partial record Pearson(Expression[] Arguments) : Function
{
    // PEARSON is the same product-moment correlation coefficient as CORREL.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        BivariateMoments.Correlation(Arguments, context);
}

[MemoryPackable]
public sealed partial record CovarianceP(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // COVARIANCE.P(array1, array2) — population covariance: Σ(x−x̄)(y−ȳ)/n; empty → #DIV/0!.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (BivariateMoments.Gather(arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return m.Count == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(m.SumXY / m.Count);
    }
}

[MemoryPackable]
public sealed partial record CovarianceS(Expression[] Arguments) : Function
{
    // COVARIANCE.S(array1, array2) — sample covariance: Σ(x−x̄)(y−ȳ)/(n−1); n < 2 → #DIV/0!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (BivariateMoments.Gather(Arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return m.Count < 2
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(m.SumXY / (m.Count - 1));
    }
}

[MemoryPackable]
public sealed partial record Rsq(Expression[] Arguments) : Function
{
    // RSQ(known_ys, known_xs) — the square of the Pearson coefficient.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (BivariateMoments.Gather(Arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (m.Count == 0 || m.SumXX == 0 || m.SumYY == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var r = m.SumXY / Math.Sqrt(m.SumXX * m.SumYY);

        return ComputedValue.Number(r * r);
    }
}

[MemoryPackable]
public sealed partial record Slope(Expression[] Arguments) : Function
{
    // SLOPE(known_ys, known_xs) — least-squares slope Σ(x−x̄)(y−ȳ)/Σ(x−x̄)²; var(x)=0 → #DIV/0!.
    // NOTE the argument order: ys first, then xs (SumXX/SumXY below are named after the
    // REGRESSOR x = second argument).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (BivariateMoments.GatherRegression(Arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return m.Count == 0 || m.SumXX == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(m.SumXY / m.SumXX);
    }
}

[MemoryPackable]
public sealed partial record Intercept(Expression[] Arguments) : Function
{
    // INTERCEPT(known_ys, known_xs) — ȳ − slope·x̄.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (BivariateMoments.GatherRegression(Arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return m.Count == 0 || m.SumXX == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(m.MeanY - m.SumXY / m.SumXX * m.MeanX);
    }
}

[MemoryPackable]
public sealed partial record Steyx(Expression[] Arguments) : Function
{
    // STEYX(known_ys, known_xs) — standard error of the predicted y:
    // sqrt[(Σ(y−ȳ)² − (Σ(x−x̄)(y−ȳ))²/Σ(x−x̄)²) / (n−2)]; fewer than 3 points → #DIV/0!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (BivariateMoments.GatherRegression(Arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (m.Count < 3 || m.SumXX == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var residual = m.SumYY - m.SumXY * m.SumXY / m.SumXX;

        return ComputedValue.Number(Math.Sqrt(residual / (m.Count - 2)));
    }
}

[MemoryPackable]
public sealed partial record ForecastLinear(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // FORECAST.LINEAR(x, known_ys, known_xs) — intercept + slope·x. Argument order: the new x
    // FIRST, then ys, then xs.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (arguments[0].Evaluate(context).CoerceToNumber(out var x) is { } xError)
        {
            return ComputedValue.Error(xError);
        }

        if (BivariateMoments.GatherRegression(arguments[1..], context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (m.Count == 0 || m.SumXX == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var slope = m.SumXY / m.SumXX;

        return ComputedValue.Number(m.MeanY - slope * m.MeanX + slope * x);
    }
}

/// <summary>Centered second moments of a pair collection: Σ(x−x̄)², Σ(y−ȳ)², Σ(x−x̄)(y−ȳ).</summary>
file readonly record struct Moments(
    int Count,
    double MeanX,
    double MeanY,
    double SumXX,
    double SumYY,
    double SumXY
);

file static class BivariateMoments
{
    /// <summary>Pairs arguments[0] as x and arguments[1] as y (CORREL/COVARIANCE order).</summary>
    public static Error? Gather(
        Expression[] arguments,
        EvaluationContext context,
        out Moments moments
    ) => Compute(arguments[0], arguments[1], context, out moments);

    /// <summary>Pairs arguments[0] as known_ys and arguments[1] as known_xs — the regression
    /// functions take ys FIRST (SLOPE/INTERCEPT/RSQ/STEYX/FORECAST).</summary>
    public static Error? GatherRegression(
        Expression[] arguments,
        EvaluationContext context,
        out Moments moments
    ) => Compute(arguments[1], arguments[0], context, out moments);

    private static Error? Compute(
        Expression xArgument,
        Expression yArgument,
        EvaluationContext context,
        out Moments moments
    )
    {
        moments = default;

        if (
            PairwiseRanges.Expand(
                xArgument,
                yArgument,
                context,
                Error.NA,
                PairwisePolicy.IgnorePair,
                out var pairs
            ) is
            { } error
        )
        {
            return error;
        }

        var n = pairs.Count;

        if (n == 0)
        {
            moments = new Moments(0, 0, 0, 0, 0, 0);
            return null;
        }

        var meanX = 0.0;
        var meanY = 0.0;

        foreach (var (x, y) in pairs)
        {
            meanX += x;
            meanY += y;
        }

        meanX /= n;
        meanY /= n;

        var sumXX = 0.0;
        var sumYY = 0.0;
        var sumXY = 0.0;

        foreach (var (x, y) in pairs)
        {
            var dx = x - meanX;
            var dy = y - meanY;

            sumXX += dx * dx;
            sumYY += dy * dy;
            sumXY += dx * dy;
        }

        moments = new Moments(n, meanX, meanY, sumXX, sumYY, sumXY);
        return null;
    }

    public static ComputedValue Correlation(Expression[] arguments, EvaluationContext context)
    {
        if (Gather(arguments, context, out var m) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return m.Count == 0 || m.SumXX == 0 || m.SumYY == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(m.SumXY / Math.Sqrt(m.SumXX * m.SumYY));
    }
}
