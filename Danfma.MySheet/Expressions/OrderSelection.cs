using System.Buffers;

namespace Danfma.MySheet.Expressions;

/// <summary>Shared argument plumbing of the order statistics: an array argument plus one numeric
/// scalar, with the array sorted ascending.
///
/// <para>This lives in its own file, <c>internal</c> rather than the local <c>file static class</c> idiom,
/// because the bounded-heap k-th selection (<see cref="KthValueStreaming"/>) and the bounds contract it
/// shares with the collected path (<see cref="KthOfSorted"/>) are used beyond OrderStatistics.cs, and a
/// file-local type is invisible across files. A second copy of the heap would let the two drift on the
/// empty / k &lt; 1 / k &gt; n edges.</para></summary>
internal static class OrderSelection
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
            && ArrayEvaluation.IsArrayEligible(arguments[0], context)
            && ArrayEvaluation.TryEvaluateStream(arguments[0], context, out var stream)
        )
        {
            return KthValueStreaming(stream, arguments[1], context, largest);
        }

        if (SortedArrayAndScalar(arguments, context, out var sorted, out var k) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return KthOfSorted(sorted, k, largest);
    }

    /// <summary>The k-th value of an ASCENDING-sorted population: SMALL's <c>sorted[k - 1]</c> and LARGE's
    /// <c>sorted[^k]</c> plus the bounds contract they share — k is truncated toward zero (Excel takes the
    /// integer part) and an empty population, k &lt; 1 or k past the count is #NUM!. Shared so every
    /// COLLECTED (non-streaming) caller reports those edges identically to the streaming path below.</summary>
    public static ComputedValue KthOfSorted(IReadOnlyList<double> sorted, double k, bool largest)
    {
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
