namespace Danfma.MySheet.Expressions;

/// <summary>
/// Shared numeric machinery of the descriptive statistics: gathering values (SUM-style referenced
/// semantics via <see cref="NumericAggregation"/>, in scan order — MODE's tie-break depends on it),
/// Excel's two percentile interpolations (PERCENTILE.INC at <c>k·(n−1)</c>, PERCENTILE.EXC at
/// <c>k·(n+1)</c>), and the variance folds reused by STDEV*/VAR*/SUBTOTAL.
/// </summary>
internal static class StatisticsMath
{
    /// <summary>Gathers the numeric values of the arguments in scan order (SUM semantics: referenced
    /// text/logicals/blanks ignored). Returns the first cell error to propagate, or <c>null</c>.</summary>
    public static Error? Collect(
        Expression[] arguments,
        EvaluationContext context,
        out List<double> values
    )
    {
        var fold = new NumericListFold { Values = [] };
        var error = NumericAggregation.Fold(arguments, context, ref fold);

        values = fold.Values;
        return error;
    }

    /// <summary>Like <see cref="Collect"/> but with the A-variant rule (AVERAGEA/STDEVA/…):
    /// referenced text counts as 0 and referenced logicals as 1/0; blanks are still ignored.</summary>
    public static Error? CollectA(
        Expression[] arguments,
        EvaluationContext context,
        out List<double> values
    )
    {
        var fold = new NumericListFold { Values = [] };
        var error = NumericAggregation.FoldA(arguments, context, ref fold);

        values = fold.Values;
        return error;
    }

    public static double Mean(List<double> values)
    {
        var total = 0.0;

        foreach (var value in values)
        {
            total += value;
        }

        return total / values.Count;
    }

    /// <summary>Σ(x−x̄)² — the DEVSQ quantity, shared by the variance folds.</summary>
    public static double SumSquaredDeviations(List<double> values)
    {
        var mean = Mean(values);
        var total = 0.0;

        foreach (var value in values)
        {
            var deviation = value - mean;
            total += deviation * deviation;
        }

        return total;
    }

    /// <summary>Sample variance (n−1 denominator); fewer than 2 values → <c>#DIV/0!</c>.</summary>
    public static Error? SampleVariance(List<double> values, out double variance)
    {
        if (values.Count < 2)
        {
            variance = 0;
            return Error.DivZero;
        }

        variance = SumSquaredDeviations(values) / (values.Count - 1);
        return null;
    }

    /// <summary>Population variance (n denominator); no values → <c>#DIV/0!</c>.</summary>
    public static Error? PopulationVariance(List<double> values, out double variance)
    {
        if (values.Count == 0)
        {
            variance = 0;
            return Error.DivZero;
        }

        variance = SumSquaredDeviations(values) / values.Count;
        return null;
    }

    /// <summary>
    /// PERCENTILE.INC over an ascending sorted list: linear interpolation at position <c>k·(n−1)</c>
    /// (0-based). Empty list or <c>k</c> outside <c>[0, 1]</c> → <c>#NUM!</c>.
    /// </summary>
    public static Error? PercentileInclusive(List<double> sorted, double k, out double result)
    {
        result = 0;

        if (sorted.Count == 0 || k is < 0 or > 1)
        {
            return Error.Num;
        }

        var position = k * (sorted.Count - 1);
        var lower = (int)Math.Floor(position);
        var fraction = position - lower;

        result = fraction == 0
            ? sorted[lower]
            : sorted[lower] + fraction * (sorted[lower + 1] - sorted[lower]);

        return null;
    }

    /// <summary>
    /// PERCENTILE.EXC over an ascending sorted list: linear interpolation at 1-based rank
    /// <c>k·(n+1)</c>. Empty list, <c>k</c> outside <c>(0, 1)</c>, or a rank outside <c>[1, n]</c>
    /// (unreachable percentile) → <c>#NUM!</c>.
    /// </summary>
    public static Error? PercentileExclusive(List<double> sorted, double k, out double result)
    {
        result = 0;

        if (sorted.Count == 0 || k is <= 0 or >= 1)
        {
            return Error.Num;
        }

        var rank = k * (sorted.Count + 1);

        if (rank < 1 || rank > sorted.Count)
        {
            return Error.Num;
        }

        var lower = (int)Math.Floor(rank);
        var fraction = rank - lower;

        result = fraction == 0
            ? sorted[lower - 1]
            : sorted[lower - 1] + fraction * (sorted[lower] - sorted[lower - 1]);

        return null;
    }
}
