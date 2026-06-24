using MemoryPack;

namespace MySheet.Expressions;

[MemoryPackable]
public sealed partial record Index(Expression[] Arguments) : Function
{
    public override object? Compute(Workbook workbook)
    {
        if (Arguments[0] is not RangeReference range)
        {
            return ErrorValue.Reference;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(workbook), out var first) is { } firstError)
        {
            return firstError;
        }

        double row;
        double column;

        if (Arguments.Length == 3)
        {
            if (ValueCoercion.TryToNumber(Arguments[2].Compute(workbook), out column) is { } columnError)
            {
                return columnError;
            }

            row = first;
        }
        else if (range.RowCount == 1)
        {
            // A single-row range takes the lone index as a column.
            row = 1;
            column = first;
        }
        else
        {
            row = first;
            column = 1;
        }

        if (row < 1 || column < 1 || row > range.RowCount || column > range.ColumnCount)
        {
            return ErrorValue.Reference;
        }

        return range.CellAt(workbook, (int)row, (int)column).Compute(workbook);
    }
}
