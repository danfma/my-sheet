using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record VLookup(Expression[] Arguments) : Function
{
    // VLOOKUP(lookup, table, column_index, [range_lookup]) — searches the table's first column.
    public override object? Compute(EvaluationContext context)
    {
        if (Arguments[1] is not RangeReference table)
        {
            return ErrorValue.Reference;
        }

        var lookup = Arguments[0].Compute(context);

        if (ValueCoercion.TryToNumber(Arguments[2].Compute(context), out var columnIndex) is { } columnError)
        {
            return columnError;
        }

        if (columnIndex < 1 || columnIndex > table.ColumnCount)
        {
            return ErrorValue.Reference;
        }

        var approximate = true;

        if (Arguments.Length == 4 &&
            ValueCoercion.TryToBool(Arguments[3].Compute(context), out approximate) is { } modeError)
        {
            return modeError;
        }

        var matchRow = -1;

        if (approximate)
        {
            ValueCoercion.TryToNumber(lookup, out var target);

            for (var row = 1; row <= table.RowCount; row++)
            {
                if (table.CellValueAt(context, row, 1) is double key && key <= target)
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
            ? table.CellValueAt(context, matchRow, (int)columnIndex)
            : ErrorValue.NotAvailable;
    }
}
