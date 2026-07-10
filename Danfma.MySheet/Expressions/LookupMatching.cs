namespace Danfma.MySheet.Expressions;

/// <summary>
/// The shared match engine behind XLOOKUP, XMATCH and LOOKUP (extracted from XLOOKUP so the mode
/// semantics stay identical): exact match first, in the chosen direction, then the mode-specific
/// fallback. <c>matchMode</c>: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
/// Returns the 0-based index of the match, or -1 when there is none.
/// </summary>
internal static class LookupMatching
{
    public static int FindMatch(
        in ComputedValue lookup,
        IReadOnlyList<ComputedValue> array,
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

    private static int Wildcard(
        in ComputedValue lookup,
        IReadOnlyList<ComputedValue> array,
        int count,
        bool reverse
    )
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

    private static int Closest(
        in ComputedValue lookup,
        IReadOnlyList<ComputedValue> array,
        int count,
        bool below
    )
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

            if (
                below
                    ? ValueCoercion.Compare(value, lookup) > 0
                    : ValueCoercion.Compare(value, lookup) < 0
            )
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
