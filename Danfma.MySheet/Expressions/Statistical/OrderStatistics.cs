using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

// The order/position statistics: sorted-copy machinery lives in StatisticsMath (the shared
// PERCENTILE.INC/.EXC interpolations); the records here add each function's argument handling and
// documented error contract. The internal static Compute methods are reused by the Compatibility
// aliases (MODE/RANK/PERCENTILE/PERCENTRANK/QUARTILE), which are distinct nodes so the un-parse
// keeps the legacy spelling.

[MemoryPackable]
public sealed partial record Median(Expression[] Arguments) : Function
{
    // MEDIAN(number1, …) — middle value (mean of the two middle values for an even count).
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

        values.Sort();

        var middle = values.Count / 2;

        return ComputedValue.Number(
            values.Count % 2 == 1 ? values[middle] : (values[middle - 1] + values[middle]) / 2
        );
    }
}

[MemoryPackable]
public sealed partial record ModeSngl(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // MODE.SNGL(number1, …) — most frequent value; ties resolve to the FIRST value encountered in
    // scan order; no value repeats → #N/A.
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (StatisticsMath.Collect(arguments, context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var counts = new Dictionary<double, int>();
        var bestValue = 0.0;
        var bestCount = 0;

        foreach (var value in values)
        {
            counts.TryGetValue(value, out var count);
            counts[value] = ++count;

            // Strict '>' keeps the FIRST value reaching the winning count (scan-order tie-break).
            if (count > bestCount)
            {
                bestCount = count;
                bestValue = value;
            }
        }

        return bestCount < 2 ? ComputedValue.Error(Error.NA) : ComputedValue.Number(bestValue);
    }
}

[MemoryPackable]
public sealed partial record Large(Expression[] Arguments) : Function
{
    // LARGE(array, k) — k-th largest; empty array, k ≤ 0 or k > n → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        OrderSelection.KthValue(Arguments, context, largest: true);
}

[MemoryPackable]
public sealed partial record Small(Expression[] Arguments) : Function
{
    // SMALL(array, k) — k-th smallest; empty array, k ≤ 0 or k > n → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        OrderSelection.KthValue(Arguments, context, largest: false);
}

[MemoryPackable]
public sealed partial record RankEq(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context, average: false);

    // RANK.EQ/RANK.AVG(number, ref, [order]) — order omitted/0 → descending, anything else →
    // ascending. Ties: EQ gives the whole group the top rank of the group; AVG gives the mean of
    // the ranks the group occupies. A number absent from ref → #N/A.
    internal static ComputedValue Compute(
        Expression[] arguments,
        EvaluationContext context,
        bool average
    )
    {
        if (arguments[0].Evaluate(context).CoerceToNumber(out var number) is { } numberError)
        {
            return ComputedValue.Error(numberError);
        }

        var ascending = false;

        if (arguments.Length == 3)
        {
            if (arguments[2].Evaluate(context).CoerceToNumber(out var order) is { } orderError)
            {
                return ComputedValue.Error(orderError);
            }

            ascending = order != 0;
        }

        if (StatisticsMath.Collect([arguments[1]], context, out var values) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var equal = 0;
        var outranking = 0;

        foreach (var value in values)
        {
            if (value == number)
            {
                equal++;
            }
            else if (ascending ? value < number : value > number)
            {
                outranking++;
            }
        }

        if (equal == 0)
        {
            return ComputedValue.Error(Error.NA);
        }

        // The tied group occupies ranks (outranking+1) … (outranking+equal).
        return ComputedValue.Number(
            average ? outranking + (equal + 1) / 2.0 : outranking + 1
        );
    }
}

[MemoryPackable]
public sealed partial record RankAvg(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        RankEq.Compute(Arguments, context, average: true);
}

[MemoryPackable]
public sealed partial record PercentileInc(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // PERCENTILE.INC(array, k) — k ∈ [0, 1], linear interpolation at k·(n−1).
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (OrderSelection.SortedArrayAndScalar(arguments, context, out var sorted, out var k)
            is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.PercentileInclusive(sorted, k, out var result) is { } numError
            ? ComputedValue.Error(numError)
            : ComputedValue.Number(result);
    }
}

[MemoryPackable]
public sealed partial record PercentileExc(Expression[] Arguments) : Function
{
    // PERCENTILE.EXC(array, k) — k ∈ (0, 1), interpolation at rank k·(n+1); unreachable → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (OrderSelection.SortedArrayAndScalar(Arguments, context, out var sorted, out var k)
            is { } error)
        {
            return ComputedValue.Error(error);
        }

        return StatisticsMath.PercentileExclusive(sorted, k, out var result) is { } numError
            ? ComputedValue.Error(numError)
            : ComputedValue.Number(result);
    }
}

[MemoryPackable]
public sealed partial record PercentRankInc(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context, exclusive: false);

    // PERCENTRANK.INC/.EXC(array, x, [significance]) — the rank of x as a fraction of the data:
    // INC ranks a data value as (count below)/(n−1), EXC as (count below + 1)/(n+1); an x between
    // two data values interpolates linearly between their ranks. The result is TRUNCATED to
    // `significance` digits (default 3; < 1 → #NUM!). x outside [min, max] → #N/A.
    internal static ComputedValue Compute(
        Expression[] arguments,
        EvaluationContext context,
        bool exclusive
    )
    {
        if (OrderSelection.SortedArrayAndScalar(arguments, context, out var sorted, out var x)
            is { } error)
        {
            return ComputedValue.Error(error);
        }

        var significance = 3.0;

        if (arguments.Length == 3)
        {
            if (arguments[2].Evaluate(context).CoerceToNumber(out significance) is { } sigError)
            {
                return ComputedValue.Error(sigError);
            }

            significance = Math.Truncate(significance);

            if (significance < 1)
            {
                return ComputedValue.Error(Error.Num);
            }
        }

        if (sorted.Count == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        if (x < sorted[0] || x > sorted[^1])
        {
            return ComputedValue.Error(Error.NA);
        }

        var fraction = InterpolatedRank(sorted, x, exclusive);

        return ComputedValue.Number(ExcelMath.TruncateToDigits(fraction, significance));
    }

    private static double InterpolatedRank(List<double> sorted, double x, bool exclusive)
    {
        // Exact hit: the documented rank of the data value itself.
        var below = 0;

        foreach (var value in sorted)
        {
            if (value < x)
            {
                below++;
            }
            else if (value == x)
            {
                return RankOf(below, sorted.Count, exclusive);
            }
        }

        // x sits strictly between sorted[below-1] and sorted[below]: interpolate between the two
        // neighboring data ranks.
        var lowerValue = sorted[below - 1];
        var upperValue = sorted[below];
        var lowerRank = RankOf(below - 1, sorted.Count, exclusive);
        var upperRank = RankOf(below, sorted.Count, exclusive);

        return lowerRank + (x - lowerValue) / (upperValue - lowerValue) * (upperRank - lowerRank);
    }

    private static double RankOf(int countBelow, int count, bool exclusive) =>
        exclusive
            ? (countBelow + 1) / (double)(count + 1)
            : countBelow / (double)Math.Max(count - 1, 1);
}

[MemoryPackable]
public sealed partial record PercentRankExc(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        PercentRankInc.Compute(Arguments, context, exclusive: true);
}

[MemoryPackable]
public sealed partial record QuartileInc(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        Compute(Arguments, context);

    // QUARTILE.INC(array, quart) — quart 0-4 (truncated) → PERCENTILE.INC(array, quart/4).
    internal static ComputedValue Compute(Expression[] arguments, EvaluationContext context)
    {
        if (OrderSelection.SortedArrayAndScalar(arguments, context, out var sorted, out var quart)
            is { } error)
        {
            return ComputedValue.Error(error);
        }

        quart = Math.Truncate(quart);

        if (quart is < 0 or > 4)
        {
            return ComputedValue.Error(Error.Num);
        }

        return StatisticsMath.PercentileInclusive(sorted, quart / 4, out var result) is { } numError
            ? ComputedValue.Error(numError)
            : ComputedValue.Number(result);
    }
}

[MemoryPackable]
public sealed partial record QuartileExc(Expression[] Arguments) : Function
{
    // QUARTILE.EXC(array, quart) — quart truncated → PERCENTILE.EXC(array, quart/4); quart 0 and 4
    // are unreachable in the exclusive definition → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (OrderSelection.SortedArrayAndScalar(Arguments, context, out var sorted, out var quart)
            is { } error)
        {
            return ComputedValue.Error(error);
        }

        quart = Math.Truncate(quart);

        return StatisticsMath.PercentileExclusive(sorted, quart / 4, out var result) is { } numError
            ? ComputedValue.Error(numError)
            : ComputedValue.Number(result);
    }
}

[MemoryPackable]
public sealed partial record TrimMean(Expression[] Arguments) : Function
{
    // TRIMMEAN(array, percent) — percent ∈ [0, 1); INT(n·percent/2) values are cut from EACH end
    // of the sorted data before averaging.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (OrderSelection.SortedArrayAndScalar(Arguments, context, out var sorted, out var percent)
            is { } error)
        {
            return ComputedValue.Error(error);
        }

        if (percent < 0 || percent >= 1 || sorted.Count == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var trim = (int)(sorted.Count * percent / 2);
        var total = 0.0;

        for (var i = trim; i < sorted.Count - trim; i++)
        {
            total += sorted[i];
        }

        return ComputedValue.Number(total / (sorted.Count - 2 * trim));
    }
}

/// <summary>Shared argument plumbing of the order statistics: an array argument plus one numeric
/// scalar, with the array sorted ascending.</summary>
file static class OrderSelection
{
    public static Error? SortedArrayAndScalar(
        Expression[] arguments,
        EvaluationContext context,
        out List<double> sorted,
        out double scalar
    )
    {
        scalar = 0;

        if (StatisticsMath.Collect([arguments[0]], context, out sorted) is { } error)
        {
            return error;
        }

        if (arguments[1].Evaluate(context).CoerceToNumber(out scalar) is { } scalarError)
        {
            return scalarError;
        }

        sorted.Sort();
        return null;
    }

    public static ComputedValue KthValue(
        Expression[] arguments,
        EvaluationContext context,
        bool largest
    )
    {
        if (SortedArrayAndScalar(arguments, context, out var sorted, out var k) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var position = (int)Math.Truncate(k);

        if (sorted.Count == 0 || position < 1 || position > sorted.Count)
        {
            return ComputedValue.Error(Error.Num);
        }

        return ComputedValue.Number(largest ? sorted[^position] : sorted[position - 1]);
    }
}
