using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record VLookup(Expression[] Arguments) : Function
{
    // VLOOKUP(lookup, table, column_index, [range_lookup]) — searches the table's first column.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[1] is not RangeReference table)
        {
            return ComputedValue.Error(Error.Ref);
        }

        var lookup = Arguments[0].Compute(context);

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var columnIndex) is { } columnError)
        {
            return ComputedValue.Error(columnError);
        }

        if (columnIndex < 1 || columnIndex > table.ColumnCount)
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

        if (lookup is ErrorValue lookupError)
        {
            return ComputedValue.From(lookupError);
        }

        var matchRow = -1;

        if (approximate)
        {
            // Largest first-column key <= lookup, assuming the table is sorted ascending. Cross-type
            // ordering (ValueCoercion.Compare) lets text keys sort lexicographically, exactly like the
            // <= operator — not only numeric keys.
            for (var row = 1; row <= table.RowCount; row++)
            {
                var key = table.CellValueAt(context, row, 1);
                if (key is null or ErrorValue)
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
                if (ValueCoercion.AreEqual(table.CellValueAt(context, row, 1), lookup))
                {
                    matchRow = row;
                    break;
                }
            }
        }

        return matchRow >= 1
            ? ComputedValue.From(table.CellValueAt(context, matchRow, (int)columnIndex))
            : ComputedValue.Error(Error.NA);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
