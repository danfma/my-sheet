using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record XLookup(Expression[] Arguments) : Function
{
    // XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
    // match_mode: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first-to-last, -1 last-to-first (binary modes not supported).
    // The match engine itself is shared with XMATCH and LOOKUP (see LookupMatching).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet lookup/return array is a structural #REF! — distinct from an empty array over an
        // existing sheet, which stays #N/A. Guard before enumerating so it is not swallowed as empty.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var lookup = Arguments[0].Evaluate(context);
        var lookupArray = ArgumentFlattening.ExpandCached(Arguments[1], context, out var lookupSnapshot);
        var returnArray = ArgumentFlattening.ExpandCached(Arguments[2], context, out _);

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

        // Forward exact (the default) → O(1) via the value→first-position hash, but only when the whole
        // lookup array is covered by the return array (so the hashed position is a valid return index — the
        // linear engine only scans the shared [0, count) prefix). Every other case keeps LookupMatching.
        if (
            (int)matchMode == 0
            && searchMode >= 0
            && lookupSnapshot is not null
            && lookupSnapshot.Count <= returnArray.Count
        )
        {
            switch (lookupSnapshot.TryExactPosition(lookup, out var hashPosition))
            {
                case ExactMatchOutcome.Found:
                    return returnArray[hashPosition - 1];
                case ExactMatchOutcome.NotFound:
                    return Arguments.Length >= 4 && Arguments[3] is not BlankValue
                        ? Arguments[3].Evaluate(context)
                        : ComputedValue.Error(Error.NA);
            }
        }

        var match = LookupMatching.FindMatch(
            lookup,
            lookupArray,
            count,
            (int)matchMode,
            reverse: searchMode < 0
        );

        if (match >= 0)
        {
            return returnArray[match];
        }

        return Arguments.Length >= 4 && Arguments[3] is not BlankValue
            ? Arguments[3].Evaluate(context)
            : ComputedValue.Error(Error.NA);
    }
}
