using System.Text.RegularExpressions;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// The shared match engine behind XLOOKUP, XMATCH and LOOKUP (extracted from XLOOKUP so the mode
/// semantics stay identical): exact match first, in the chosen direction, then the mode-specific
/// fallback. <c>matchMode</c>: 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
/// Returns the 0-based index of the match, or -1 when there is none.
/// </summary>
internal static class LookupMatching
{
    public static ComputedValue EvaluateKey(
        Expression expression,
        EvaluationContext context,
        out bool absent,
        out Reference? resolvedReference
    )
    {
        if (
            ResolvedReferenceValue.TryClassify(
                expression,
                context,
                out var reference,
                out var value,
                out absent,
                out var unresolvedValue
            )
        )
        {
            resolvedReference = reference;
            return expression is Reference && value.Kind == ComputedValueKind.Reference
                ? ComputedValue.Error(Error.Value)
                : value;
        }

        resolvedReference = null;
        var evaluated = unresolvedValue ?? expression.Evaluate(context);
        absent = expression switch
        {
            AnchoredCellReference anchored => IsAbsent(anchored, context),
            NameReference name => context.IsAbsentName(name.Name),
            _ => false,
        };
        return evaluated;
    }

    internal static bool IsAbsent(AnchoredCellReference cell, EvaluationContext context)
    {
        if (!context.Workbook.Sheets.TryGetValue(cell.SheetName, out var sheet))
        {
            return false;
        }

        var (column, row) = cell.Effective(context);
        return !sheet.ContainsKey(new CellAddress(column, row).ToId());
    }

    internal readonly struct ExactMatcher
    {
        private readonly ComputedValue _lookup;
        private readonly Regex? _wildcard;
        private readonly bool _blankOnly;
        private readonly string? _exactText;

        public ExactMatcher(in ComputedValue lookup, bool blankOnly = false)
        {
            _lookup = lookup;
            _blankOnly = blankOnly;
            _exactText = lookup.TryGetText(out var text) && text.Length == 0 ? text : null;
            _wildcard =
                UsesWildcards(lookup) && lookup.TryGetText(out var pattern)
                    ? Criteria.BuildWildcardRegex(pattern)
                    : null;
        }

        public bool Matches(in ComputedValue candidate) =>
            _blankOnly ? candidate.Kind == ComputedValueKind.Blank
            : _exactText is not null ? IsExactText(candidate, _exactText)
            : _wildcard is null ? ValueCoercion.AreEqual(candidate, _lookup)
            : candidate.TryGetText(out var text) && IsWildcardMatch(_wildcard, text);
    }

    public static bool UsesWildcards(in ComputedValue lookup) =>
        lookup.TryGetText(out var pattern)
        && (pattern.Contains('*') || pattern.Contains('?') || pattern.Contains('~'));

    public static ExactMatcher TableExactMatcher(in ComputedValue lookup) => new(lookup);

    public static bool IsExactText(in ComputedValue candidate, string lookupText) =>
        candidate.TryGetText(out var candidateText)
        && string.Equals(candidateText, lookupText, StringComparison.OrdinalIgnoreCase);

    public static int FindMatch(
        in ComputedValue lookup,
        IReadOnlyList<ComputedValue> array,
        int count,
        int matchMode,
        bool reverse,
        bool absentMatchesOnlyBlank = false
    )
    {
        // Exact match first (in the chosen direction) for every mode except wildcard.
        if (matchMode != 2)
        {
            var matcher = new ExactMatcher(lookup, absentMatchesOnlyBlank);
            for (var k = 0; k < count; k++)
            {
                var i = reverse ? count - 1 - k : k;
                var candidate = array[i];
                if (
                    absentMatchesOnlyBlank
                        ? reverse
                            ? candidate.TryGetText(out var text) && text.Length == 0
                            : candidate.Kind == ComputedValueKind.Blank
                        : matcher.Matches(candidate)
                )
                {
                    return i;
                }
            }
        }

        if (matchMode is -1 or 1 && lookup.TryGetText(out var exactText) && exactText.Length == 0)
        {
            return -1;
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
        var matcher = TableExactMatcher(lookup);

        for (var k = 0; k < count; k++)
        {
            var i = reverse ? count - 1 - k : k;
            if (matcher.Matches(array[i]))
            {
                return i;
            }
        }

        return -1;
    }

    // A timeout makes only that candidate a non-match; lookup scans have no error channel for it.
    private static bool IsWildcardMatch(Regex regex, string text)
    {
        try
        {
            return regex.IsMatch(text);
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }
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
