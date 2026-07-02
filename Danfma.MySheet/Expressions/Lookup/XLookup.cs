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
