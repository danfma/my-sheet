using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

[MemoryPackable]
public sealed partial record Index(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // Mini-CSE array-argument forms (the array is a computed vector, not a reference):
        //   • ROW of a whole/one-sided-open column is the IDENTITY row vector [top, top+1, …]; INDEX(…,n)
        //     returns its n-th worksheet row number WITHOUT materializing the grid-less column (the K1
        //     "INDEX(ROW($A:$A), SMALL(…))" idiom). Documented special case for Row([OpenRangeReference]).
        if (Arguments[0] is Row { Arguments: [OpenRangeReference openColumn] })
        {
            return IndexIntoOpenRowNumbers(openColumn, context);
        }

        //   • Any other array-eligible, non-reference first argument — ROW(B2:B5), IF(range=…,…) — is
        //     materialized row-major and indexed. References fall through to the concrete-range path below,
        //     so plain INDEX(A1:C10, r, c) is untouched.
        if (
            Arguments[0] is not Reference
            && ArrayEvaluation.IsArrayEligible(Arguments[0])
            && ArrayEvaluation.TryEvaluateStream(Arguments[0], context, out var array)
        )
        {
            return IndexIntoArray(array, context);
        }

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

    // Indexes an element-wise vector row-major (the mini-CSE array form) through the LAZY stream: only the
    // selected element is computed (no ComputedValue[] materialized). Mirrors the concrete-range branch's
    // row/column resolution: a 3-arg call takes (row, column); a 2-arg call maps the lone index to the array's
    // only axis (column for a single row, row otherwise). Out of bounds → #REF!.
    private ComputedValue IndexIntoArray(ArrayEvaluation.ArrayStream array, EvaluationContext context)
    {
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
        else if (array.Rows == 1)
        {
            row = 1;
            column = first;
        }
        else
        {
            row = first;
            column = 1;
        }

        if (row < 1 || column < 1 || row > array.Rows || column > array.Columns)
        {
            return ComputedValue.Error(Error.Ref);
        }

        // Row-major layout (ArrayEvaluation lays out row-then-column): (r,c) 1-based → (r-1)·Columns + (c-1).
        return array.ElementAt(((int)row - 1) * array.Columns + ((int)column - 1));
    }

    // ROW of an open column is the identity vector [top, top+1, …] with top = RowMin (or 1 when the top is
    // open). INDEX(ROW($A:$A), n) therefore returns top+n-1 directly — no column is materialized, honoring
    // the grid-less model (an open bottom has no fixed 1,048,576 cap). n < 1, or n past a bounded bottom,
    // is #REF!. A 3-arg form must select the single column (column 1).
    private ComputedValue IndexIntoOpenRowNumbers(OpenRangeReference open, EvaluationContext context)
    {
        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is { } firstError)
        {
            return ComputedValue.Error(firstError);
        }

        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out var column) is { } columnError)
            {
                return ComputedValue.Error(columnError);
            }

            if (column is < 1 or > 1)
            {
                return ComputedValue.Error(Error.Ref);
            }
        }

        var n = (int)Math.Truncate(first);
        var top = open.RowMin ?? 1;

        if (n < 1 || (open.RowMax is { } bottom && top + n - 1 > bottom))
        {
            return ComputedValue.Error(Error.Ref);
        }

        return ComputedValue.Number(top + n - 1);
    }

    // Mirrors the concrete-range branch of Evaluate, but yields the target CELL ADDRESS as a Reference
    // instead of reading its value. Array forms (the mini-CSE vector and the open-column ROW identity) have
    // no cell address to hand back, so they return false and fall through to normal evaluation elsewhere.
    public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
    {
        reference = null;

        // Array forms (mini-CSE vector, open-column ROW identity) have no cell address.
        if (
            Arguments[0] is Row { Arguments: [OpenRangeReference] }
            || (Arguments[0] is not Reference && ArrayEvaluation.IsArrayEligible(Arguments[0]))
        )
        {
            return false;
        }

        if (
            !NamedReferences.TryResolveReference(Arguments[0], context, out var resolved)
            || resolved is not RangeReference range
        )
        {
            return false;
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var first) is not null)
        {
            return false;
        }

        double row;
        double column;
        if (Arguments.Length == 3)
        {
            if (Arguments[2].Evaluate(context).CoerceToNumber(out column) is not null)
            {
                return false;
            }

            row = first;
        }
        else if (range.RowCount == 1)
        {
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
            return false;
        }

        // Mirrors RangeReference.CellComputedValueAt's normalization: StartId is not guaranteed to be the
        // top-left corner (e.g. A3:A1 has StartId="A3"), so the origin is the min of both corners.
        var start = CellAddress.Parse(range.StartId);
        var end = CellAddress.Parse(range.EndId);
        var originColumn = Math.Min(start.Column, end.Column);
        var originRow = Math.Min(start.Row, end.Row);
        var target = new CellAddress(originColumn + (int)column - 1, originRow + (int)row - 1);
        reference = new CellReference(target.ToId(), range.SheetName);
        return true;
    }
}
