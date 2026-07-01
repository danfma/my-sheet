using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record XLookup(Expression[] Arguments) : Function
{
    // XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
    // match_mode: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first-to-last, -1 last-to-first (binary modes not supported).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var lookup = Arguments[0].Compute(context);
        var lookupArray = ArgumentFlattening.Expand(Arguments[1], context);
        var returnArray = ArgumentFlattening.Expand(Arguments[2], context);

        var matchMode = 0.0;
        if (
            Arguments.Length >= 5
            && Arguments[4].Evaluate(context).CoerceToNumber(out matchMode) is { } matchError
        )
        {
            return ComputedValue.Error(matchError);
        }

        var searchMode = 1.0;
        if (
            Arguments.Length >= 6
            && Arguments[5].Evaluate(context).CoerceToNumber(out searchMode) is { } searchError
        )
        {
            return ComputedValue.Error(searchError);
        }

        var count = Math.Min(lookupArray.Count, returnArray.Count);
        var match = FindMatch(lookup, lookupArray, count, (int)matchMode, reverse: searchMode < 0);

        if (match >= 0)
        {
            return ComputedValue.From(returnArray[match]);
        }

        return Arguments.Length >= 4 && Arguments[3] is not BlankValue
            ? ComputedValue.From(Arguments[3].Compute(context))
            : ComputedValue.Error(Error.NA);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();

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
        // An error has no place in the ordering, so there is no closest match.
        if (lookup is ErrorValue)
        {
            return -1;
        }

        // below: exact-or-next-smaller -> largest value <= lookup. !below: exact-or-next-larger ->
        // smallest value >= lookup. Cross-type ordering (ValueCoercion.Compare) lets text keys sort
        // lexicographically, exactly like the <= operator — not only numeric keys.
        var best = -1;
        object? bestValue = null;

        for (var i = 0; i < count; i++)
        {
            var value = array[i];
            if (value is null or ErrorValue)
            {
                continue;
            }

            if (below ? ValueCoercion.Compare(value, lookup) > 0 : ValueCoercion.Compare(value, lookup) < 0)
            {
                continue;
            }

            if (
                best < 0
                || (
                    below
                        ? ValueCoercion.Compare(value, bestValue) > 0
                        : ValueCoercion.Compare(value, bestValue) < 0
                )
            )
            {
                best = i;
                bestValue = value;
            }
        }

        return best;
    }
}
