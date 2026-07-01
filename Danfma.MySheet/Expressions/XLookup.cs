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
        var lookup = Arguments[0].Evaluate(context);
        var lookupArray = ArgumentFlattening.ExpandComputedValues(Arguments[1], context);
        var returnArray = ArgumentFlattening.ExpandComputedValues(Arguments[2], context);

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
            return returnArray[match];
        }

        return Arguments.Length >= 4 && Arguments[3] is not BlankValue
            ? Arguments[3].Evaluate(context)
            : ComputedValue.Error(Error.NA);
    }

    private static int FindMatch(
        in ComputedValue lookup,
        List<ComputedValue> array,
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

    private static int Wildcard(in ComputedValue lookup, List<ComputedValue> array, int count, bool reverse)
    {
        var pattern = lookup.TryGetText(out var p) ? p : string.Empty;

        for (var k = 0; k < count; k++)
        {
            var i = reverse ? count - 1 - k : k;
            if (array[i].TryGetText(out var text) && Criteria.WildcardMatch(pattern, text))
            {
                return i;
            }
        }

        return -1;
    }

    private static int Closest(in ComputedValue lookup, List<ComputedValue> array, int count, bool below)
    {
        // An error has no place in the ordering, so there is no closest match.
        if (lookup.Kind == ComputedValueKind.Error)
        {
            return -1;
        }

        // below: exact-or-next-smaller -> largest value <= lookup. !below: exact-or-next-larger ->
        // smallest value >= lookup. Cross-type ordering (ValueCoercion.Compare) lets text keys sort
        // lexicographically, exactly like the <= operator — not only numeric keys.
        var best = -1;
        ComputedValue bestValue = default;

        for (var i = 0; i < count; i++)
        {
            var value = array[i];
            if (value.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
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
