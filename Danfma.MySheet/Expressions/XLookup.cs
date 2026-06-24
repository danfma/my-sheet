using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record XLookup(Expression[] Arguments) : Function
{
    // XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
    // match_mode: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first-to-last, -1 last-to-first (binary modes not supported).
    public override object? Compute(EvaluationContext context)
    {
        var lookup = Arguments[0].Compute(context);
        var lookupArray = ArgumentFlattening.Expand(Arguments[1], context);
        var returnArray = ArgumentFlattening.Expand(Arguments[2], context);

        var matchMode = 0.0;
        if (
            Arguments.Length >= 5
            && ValueCoercion.TryToNumber(Arguments[4].Compute(context), out matchMode)
                is { } matchError
        )
        {
            return matchError;
        }

        var searchMode = 1.0;
        if (
            Arguments.Length >= 6
            && ValueCoercion.TryToNumber(Arguments[5].Compute(context), out searchMode)
                is { } searchError
        )
        {
            return searchError;
        }

        var count = Math.Min(lookupArray.Count, returnArray.Count);
        var match = FindMatch(lookup, lookupArray, count, (int)matchMode, reverse: searchMode < 0);

        if (match >= 0)
        {
            return returnArray[match];
        }

        return Arguments.Length >= 4 && Arguments[3] is not BlankValue
            ? Arguments[3].Compute(context)
            : ErrorValue.NotAvailable;
    }

    private static int FindMatch(
        object? lookup,
        List<object?> array,
        int count,
        int matchMode,
        bool reverse
    )
    {
        // Exact match first (in the chosen direction) for every mode except wildcard.
        if (matchMode != 2)
        {
            for (var k = 0; k < count; k++)
            {
                var i = reverse ? count - 1 - k : k;
                if (ValueCoercion.AreEqual(array[i], lookup))
                {
                    return i;
                }
            }
        }

        return matchMode switch
        {
            2 => Wildcard(lookup, array, count, reverse),
            -1 => Closest(lookup, array, count, below: true),
            1 => Closest(lookup, array, count, below: false),
            _ => -1,
        };
    }

    private static int Wildcard(object? lookup, List<object?> array, int count, bool reverse)
    {
        var pattern = lookup as string ?? string.Empty;

        for (var k = 0; k < count; k++)
        {
            var i = reverse ? count - 1 - k : k;
            if (array[i] is string text && Criteria.WildcardMatch(pattern, text))
            {
                return i;
            }
        }

        return -1;
    }

    private static int Closest(object? lookup, List<object?> array, int count, bool below)
    {
        if (ValueCoercion.TryToNumber(lookup, out var target) is not null)
        {
            return -1;
        }

        var best = -1;
        var bestValue = below ? double.NegativeInfinity : double.PositiveInfinity;

        for (var i = 0; i < count; i++)
        {
            if (array[i] is not double value)
            {
                continue;
            }

            if (below ? value <= target && value > bestValue : value >= target && value < bestValue)
            {
                best = i;
                bestValue = value;
            }
        }

        return best;
    }
}
