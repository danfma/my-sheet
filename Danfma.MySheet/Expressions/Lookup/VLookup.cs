using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record VLookup(Expression[] Arguments) : Function
{
    // VLOOKUP(lookup, table, column_index, [range_lookup]) — searches the table's first column.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // The table may be written directly or through a defined name that stands for a range.
        if (
            !NamedReferences.TryResolveReference(Arguments[1], context, out var reference)
            || reference is not RangeReference table
        )
        {
            return ComputedValue.Error(Error.Ref);
        }

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

        if (columnIndex > table.ColumnCount)
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

        var matchRow = -1;

        if (approximate)
        {
            // Largest first-column key <= lookup, assuming the table is sorted ascending. Cross-type
            // ordering (ValueCoercion.Compare) lets text keys sort lexicographically, exactly like the
            // <= operator — not only numeric keys.
            for (var row = 1; row <= table.RowCount; row++)
            {
                var key = table.CellComputedValueAt(context, row, 1);
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
        else
        {
            for (var row = 1; row <= table.RowCount; row++)
            {
                if (ValueCoercion.AreEqual(table.CellComputedValueAt(context, row, 1), lookup))
                {
                    matchRow = row;
                    break;
                }
            }
        }

        return matchRow >= 1
            ? table.CellComputedValueAt(context, matchRow, (int)columnIndex)
            : ComputedValue.Error(Error.NA);
    }
}
