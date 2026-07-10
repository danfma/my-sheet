using System.Buffers;
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
        IReadOnlyList<double> values;

        // A single big populated range → the shared per-epoch sorted-number view (built once); otherwise the
        // ordinary collect-then-sort. The first cell error propagates exactly as StatisticsMath.Collect does.
        if (
            Arguments is [Reference reference]
            && context.Workbook.TryGetRangeSnapshot(reference, context) is { } snapshot
        )
        {
            values = snapshot.SortedNumericValues(out var firstError);

            if (firstError is { } arrayError)
            {
                return ComputedValue.Error(arrayError);
            }
        }
        else
        {
            if (StatisticsMath.Collect(Arguments, context, out var collected) is { } error)
            {
                return ComputedValue.Error(error);
            }

            collected.Sort();
            values = collected;
        }

        if (values.Count == 0)
        {
            return ComputedValue.Error(Error.Num);
        }

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

        int equal;
        int outranking;

        // A big populated range → the shared per-epoch sorted-number view (built once): two binary
        // searches derive the same equal/outranking counts the linear scan below computes cell by cell,
        // collapsing the O(n) scan (O(n²) over a whole column dragged down thousands of rows) to
        // O(log n). The first cell error propagates exactly as StatisticsMath.Collect does.
        if (
            arguments[1] is Reference reference
            && context.Workbook.TryGetRangeSnapshot(reference, context) is { } snapshot
        )
        {
            var (countLess, countEqual, countGreater) = snapshot.NumericRankCounts(
                number,
                out var arrayError
            );

            if (arrayError is { } propagated)
            {
                return ComputedValue.Error(propagated);
            }

            equal = countEqual;
            outranking = ascending ? countLess : countGreater;
        }
        else
        {
            if (StatisticsMath.Collect([arguments[1]], context, out var values) is { } error)
            {
                return ComputedValue.Error(error);
            }

            equal = 0;
            outranking = 0;

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
        }

        if (equal == 0)
        {
            return ComputedValue.Error(Error.NA);
        }

        // The tied group occupies ranks (outranking+1) … (outranking+equal).
        return ComputedValue.Number(average ? outranking + (equal + 1) / 2.0 : outranking + 1);
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
        if (
            OrderSelection.SortedArrayAndScalar(arguments, context, out var sorted, out var k) is
            { } error
        )
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
        if (
            OrderSelection.SortedArrayAndScalar(Arguments, context, out var sorted, out var k) is
            { } error
        )
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
        if (
            OrderSelection.SortedArrayAndScalar(arguments, context, out var sorted, out var x) is
            { } error
        )
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

    private static double InterpolatedRank(IReadOnlyList<double> sorted, double x, bool exclusive)
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
        if (
            OrderSelection.SortedArrayAndScalar(
                arguments,
                context,
                out var sorted,
                out var quart
            ) is
            { } error
        )
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
        if (
            OrderSelection.SortedArrayAndScalar(
                Arguments,
                context,
                out var sorted,
                out var quart
            ) is
            { } error
        )
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
        if (
            OrderSelection.SortedArrayAndScalar(
                Arguments,
                context,
                out var sorted,
                out var percent
            ) is
            { } error
        )
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
        out IReadOnlyList<double> sorted,
        out double scalar
    )
    {
        scalar = 0;

        // A big populated array argument is served from the shared per-epoch sorted-number view (built
        // once), so SMALL/LARGE/PERCENTILE/QUARTILE/PERCENTRANK/TRIMMEAN over a whole column stop
        // re-collecting and re-sorting per formula. The first cell error propagates exactly as
        // StatisticsMath.Collect would, and BEFORE the scalar is evaluated (same order as the linear path).
        if (
            arguments[0] is Reference reference
            && context.Workbook.TryGetRangeSnapshot(reference, context) is { } snapshot
        )
        {
            var numbers = snapshot.SortedNumericValues(out var firstError);
            sorted = numbers;

            if (firstError is { } arrayError)
            {
                return arrayError;
            }

            return arguments[1].Evaluate(context).CoerceToNumber(out scalar);
        }

        if (StatisticsMath.Collect([arguments[0]], context, out var collected) is { } error)
        {
            sorted = collected;
            return error;
        }

        if (arguments[1].Evaluate(context).CoerceToNumber(out scalar) is { } scalarError)
        {
            sorted = collected;
            return scalarError;
        }

        collected.Sort();
        sorted = collected;
        return null;
    }

    public static ComputedValue KthValue(
        Expression[] arguments,
        EvaluationContext context,
        bool largest
    )
    {
        // An array-eligible (mini-CSE) first argument — SMALL(IF(range=…,ROW(…)),k) and friends — is selected
        // by STREAMING through a bounded heap of k slots (O(n log k), only k slots allocated), instead of
        // materializing the vector and sorting it (the ~14MB LOH allocation the fishing pass targeted). A plain
        // reference keeps the shared sorted-view/collect path below (snapshot reuse, PERCENTILE, …).
        if (
            arguments[0] is not Reference
            && ArrayEvaluation.IsArrayEligible(arguments[0])
            && ArrayEvaluation.TryEvaluateStream(arguments[0], context, out var stream)
        )
        {
            return KthValueStreaming(stream, arguments[1], context, largest);
        }

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

    // Streaming k-th value over the mini-CSE array (no vector, no sort). Preserves the linear path's contract
    // EXACTLY: the numeric gathering is RANGE semantics (errors → first-in-scan-order propagates; logicals,
    // text and blanks are ignored) and the scan visits EVERY element so a late cell error still wins even once
    // k is satisfied; an empty gather or an out-of-range k → #NUM!. The k scalar is coerced up front so the
    // heap can be sized, but its error is DEFERRED behind the array error — matching the linear order, where
    // Collect returns the array's first error before arguments[1] is ever evaluated. Evaluating k early is
    // observationally inert (pure functional; a volatile k taints the same enclosing cell frame either way).
    private static ComputedValue KthValueStreaming(
        ArrayEvaluation.ArrayStream stream,
        Expression kArgument,
        EvaluationContext context,
        bool largest
    )
    {
        var kError = kArgument.Evaluate(context).CoerceToNumber(out var k);
        var position = kError is null ? (int)Math.Truncate(k) : 0;
        var capacity = position >= 1 ? position : 0;

        var buffer = capacity > 0 ? ArrayPool<double>.Shared.Rent(capacity) : [];

        try
        {
            // SMALL keeps the k SMALLEST (a max-heap whose root is the k-th smallest); LARGE keeps the k
            // LARGEST (a min-heap whose root is the k-th largest).
            var heap = new BoundedHeap(buffer, capacity, keepLargest: largest);
            Error? arrayError = null;
            var count = 0;

            foreach (var element in stream)
            {
                if (element.TryGetError(out var cellError))
                {
                    arrayError ??= cellError;
                }
                else if (element.TryGetNumber(out var number))
                {
                    count++;
                    heap.Offer(number);
                }

                // Referenced text, logicals and blanks are ignored — exactly as NumericAggregation does.
            }

            if (arrayError is { } propagated)
            {
                return ComputedValue.Error(propagated);
            }

            if (kError is { } scalarError)
            {
                return ComputedValue.Error(scalarError);
            }

            if (count == 0 || position < 1 || position > count)
            {
                return ComputedValue.Error(Error.Num);
            }

            return ComputedValue.Number(heap.Root);
        }
        finally
        {
            if (capacity > 0)
            {
                ArrayPool<double>.Shared.Return(buffer);
            }
        }
    }

    // A fixed-capacity binary heap over a rented span, used for bounded top-k selection. keepLargest=true keeps
    // the k largest (a MIN-heap: the root is the smallest kept, evicted when a larger value arrives);
    // keepLargest=false keeps the k smallest (a MAX-heap: the root is the largest kept). Either way the root is
    // the k-th value once k numbers have been offered.
    private struct BoundedHeap
    {
        private readonly double[] _slots;
        private readonly int _capacity;
        private readonly bool _keepLargest;
        private int _size;

        public BoundedHeap(double[] slots, int capacity, bool keepLargest)
        {
            _slots = slots;
            _capacity = capacity;
            _keepLargest = keepLargest;
            _size = 0;
        }

        // The root: the value nearest the "evict" end — the k-th value once the heap is full.
        public readonly double Root => _slots[0];

        public void Offer(double value)
        {
            if (_capacity == 0)
            {
                return;
            }

            if (_size < _capacity)
            {
                _slots[_size] = value;
                SiftUp(_size);
                _size++;
            }
            else if (HigherPriority(_slots[0], value))
            {
                // The root is more extreme (in the evict direction) than the incoming value → it should leave
                // the kept set and the incoming value takes its place.
                _slots[0] = value;
                SiftDown(0);
            }
        }

        // "Higher priority" = closer to the root = the value evicted first: the LARGEST for a max-heap
        // (keep-smallest), the SMALLEST for a min-heap (keep-largest).
        private readonly bool HigherPriority(double a, double b) => _keepLargest ? a < b : a > b;

        private void SiftUp(int index)
        {
            while (index > 0)
            {
                var parent = (index - 1) / 2;
                if (!HigherPriority(_slots[index], _slots[parent]))
                {
                    break;
                }

                (_slots[index], _slots[parent]) = (_slots[parent], _slots[index]);
                index = parent;
            }
        }

        private void SiftDown(int index)
        {
            while (true)
            {
                var left = 2 * index + 1;
                var right = left + 1;
                var top = index;

                if (left < _size && HigherPriority(_slots[left], _slots[top]))
                {
                    top = left;
                }

                if (right < _size && HigherPriority(_slots[right], _slots[top]))
                {
                    top = right;
                }

                if (top == index)
                {
                    break;
                }

                (_slots[index], _slots[top]) = (_slots[top], _slots[index]);
                index = top;
            }
        }
    }
}
