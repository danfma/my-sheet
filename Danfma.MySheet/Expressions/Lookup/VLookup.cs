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

        // The table may be written directly or through a defined name that stands for a range.
        if (
            !NamedReferences.TryResolveReference(Arguments[1], context, out var reference)
            || reference is not RangeReference table
        )
        {
            return ComputedValue.Error(Error.Ref);
        }

        // Bounds + the dense sheet handle are resolved ONCE here, not re-parsed/re-resolved on every row of the
        // linear fallback scan below. This is a pure, side-effect-free read of the table's own corners, so
        // hoisting it ahead of the argument evaluation below does not change Arguments' evaluation order.
        var workbook = context.Workbook;
        var bounds = table.GetBounds();
        var handle = workbook.ResolveDenseHandle(table.SheetName);

        var lookup = Arguments[0].Evaluate(context);

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

        if (columnIndex > bounds.ColumnCount)
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

        if (lookup.Kind == ComputedValueKind.Error)
        {
            return lookup;
        }

        // The first column is a sub-range of the table; its per-epoch snapshot serves the key search O(1)
        // (exact) / O(log n) (approximate). A 1-based snapshot position IS the 1-based table row, because the
        // key column enumerates top-to-bottom. Built LAZILY: TryGetRangeSnapshot would reject a table below the
        // cache's own size threshold anyway, so a small table (the common case) skips the RangeReference + two
        // CellAddress.ToId allocations entirely and goes straight to the linear scan.
        RangeSnapshot? keySnapshot = null;

        if (!workbook.RangeCacheDisabled && bounds.RowCount >= Workbook.RangeCacheMinimumCells)
        {
            var keyColumn = new RangeReference(
                new CellAddress(bounds.LeftColumn, bounds.TopRow).ToId(),
                new CellAddress(bounds.LeftColumn, bounds.BottomRow).ToId(),
                table.SheetName
            );
            keySnapshot = workbook.TryGetRangeSnapshot(keyColumn, context);
        }

        var matchRow = -1;

        if (approximate)
        {
            // Largest first-column key <= lookup, assuming the table is sorted ascending. Cross-type
            // ordering (ValueCoercion.Compare) lets text keys sort lexicographically, exactly like the
            // <= operator — not only numeric keys.
            if (keySnapshot is not null)
            {
                var position = keySnapshot.ApproximateAscendingPosition(lookup);
                matchRow = position >= 1 ? position : -1;
            }
            else
            {
                for (var row = 1; row <= bounds.RowCount; row++)
                {
                    var key = table.CellComputedValueAt(workbook, handle, bounds, row, 1);
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
                for (var row = 1; row <= bounds.RowCount; row++)
                {
                    if (
                        ValueCoercion.AreEqual(
                            table.CellComputedValueAt(workbook, handle, bounds, row, 1),
                            lookup
                        )
                    )
                    {
                        matchRow = row;
                        break;
                    }
                }
            }
        }

        return matchRow >= 1
            ? table.CellComputedValueAt(workbook, handle, bounds, matchRow, (int)columnIndex)
            : ComputedValue.Error(Error.NA);
    }
}
