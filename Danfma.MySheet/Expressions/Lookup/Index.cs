using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Index(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // The array may be a literal range or a defined name that stands for one.
        if (
            !NamedReferences.TryResolveReference(Arguments[0], context, out var reference)
            || reference is not RangeReference range
        )
        {
            return ComputedValue.Error(Error.Ref);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        double row;
        double column;

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out column) is { } columnError)
            {
                return ComputedValue.Error(columnError);
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
            return ComputedValue.Error(Error.Ref);
        }

        return range.CellComputedValueAt(context, (int)row, (int)column);
    }
}
