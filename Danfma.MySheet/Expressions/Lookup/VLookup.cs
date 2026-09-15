using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record VLookup(Expression[] Arguments) : Function
{
    // VLOOKUP(lookup, table, column_index, [range_lookup]) — searches the table's first column.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet table is a structural #REF! — a BOUNDED ghost range would otherwise scan its cells,
        // skip the per-cell #REF! keys, and degrade to #N/A. Guard before the table is inspected.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (!LookupGrid.TryCreate(Arguments, context, out var grid, out var tableAnswer))
        {
            return tableAnswer;
        }

        var lookup = Arguments[0].Evaluate(context);
        var approximateLookup =
            lookup.TryGetText(out var lookupText) && lookupText.Length == 0
                ? FindExactText(grid, lookup)
                : -1;

        if (ReferencePosition.IsLookupValueError(Arguments[0], lookup, context, out var valueError))
        {
            return valueError;
        }

        if (lookup.Kind == ComputedValueKind.Error)
        {
            return lookup;
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var columnIndex) is { } columnError)
        {
            return ComputedValue.Error(columnError);
        }

        // Per the docs: col_index_num < 1 -> #VALUE!, greater than the table's columns -> #REF!
        // (mirrors HLOOKUP's row_index_num rule).
        if (columnIndex < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (columnIndex > grid.Columns)
        {
            return ComputedValue.Error(Error.Ref);
        }

        var approximate = true;

        if (
            Arguments.Length == 4
            && Arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        // Sweep item 33: a zero-row table has no key to find (oracle: #N/A in both entry modes).
        if (grid.Rows == 0)
        {
            return ComputedValue.Error(Error.NA);
        }

        // The first column is a sub-range of the table; its per-epoch snapshot serves the key search O(1)
        // (exact) / O(log n) (approximate). A 1-based snapshot position IS the 1-based table row, because the
        // key column enumerates top-to-bottom. Built LAZILY: TryGetRangeSnapshot would reject a table below the
        // cache's own size threshold anyway, so a small table (the common case) skips the RangeReference + two
        // CellAddress.ToId allocations entirely and goes straight to the linear scan.
        RangeSnapshot? keySnapshot = null;

        keySnapshot = grid.TryGetKeySnapshot(context, vertical: true);

        var matchRow = -1;

        if (approximate)
        {
            if (approximateLookup >= 1)
            {
                matchRow = approximateLookup;
            }
            // Largest first-column key <= lookup, assuming the table is sorted ascending. Cross-type
            // ordering (ValueCoercion.Compare) lets text keys sort lexicographically, exactly like the
            // <= operator — not only numeric keys.
            else if (
                keySnapshot is not null
                && (lookup.Kind != ComputedValueKind.Text || !LookupMatching.UsesWildcards(lookup))
            )
            {
                var position = keySnapshot.ApproximateAscendingPosition(lookup);
                matchRow = position >= 1 ? position : -1;
            }
            else
            {
                for (var row = 1; row <= grid.Rows; row++)
                {
                    var key = grid.At(row, 1);
                    if (key.Kind is ComputedValueKind.Blank or ComputedValueKind.Error)
                    {
                        continue;
                    }

                    if (ValueCoercion.Compare(key, lookup) <= 0)
                    {
                        matchRow = row;
                    }
                }
            }
        }
        else
        {
            // Exact → the value→first-position hash; a blank-equivalent lookup falls back to the linear scan.
            if (keySnapshot is not null)
            {
                switch (keySnapshot.TryExactPosition(lookup, out var position))
                {
                    case ExactMatchOutcome.Found:
                        matchRow = position;
                        break;
                    case ExactMatchOutcome.NotFound:
                        return ComputedValue.Error(Error.NA);
                }
            }

            if (matchRow < 1)
            {
                var matches = LookupMatching.TableExactMatcher(lookup);
                for (var row = 1; row <= grid.Rows; row++)
                {
                    if (matches.Matches(grid.At(row, 1)))
                    {
                        matchRow = row;
                        break;
                    }
                }
            }
        }

        return matchRow >= 1 ? grid.At(matchRow, (int)columnIndex) : ComputedValue.Error(Error.NA);
    }

    private static int FindExactText(LookupGrid grid, in ComputedValue lookup)
    {
        lookup.TryGetText(out var lookupText);
        var match = -1;
        for (var row = 1; row <= grid.Rows; row++)
        {
            if (
                grid.At(row, 1).TryGetText(out var candidateText)
                && string.Equals(candidateText, lookupText, StringComparison.OrdinalIgnoreCase)
            )
            {
                match = row;
            }
        }

        return match;
    }
}
