namespace Danfma.MySheet.Expressions;

// ==============================================================================================
// The lazy operand tree: built once (array leaves resolved once, scalar sub-expressions evaluated once),
// then each element is computed on demand. The classes mirror the previous materialized `Operand.At`
// element-for-element, so both consumption shapes agree bit-for-bit.
// ==============================================================================================

/// <summary>
/// A recursively-built operand: either a scalar (to broadcast) or a rectangular array whose elements are
/// computed on demand by <see cref="At"/>. Building FAILS (returns <c>false</c> from
/// <see cref="ArrayEvaluation.TryBuildOperand"/>) only when the sub-tree reaches an open/whole-column
/// range in an array position — the cost guard — so the whole evaluation degrades to "not an array".
/// </summary>
internal abstract class ArrayOperand
{
    public abstract bool IsArray { get; }
    public abstract int Rows { get; }
    public abstract int Columns { get; }

    /// <summary>The pre-evaluated scalar (valid only when <see cref="IsArray"/> is <c>false</c>).</summary>
    public virtual ComputedValue Scalar => default;

    /// <summary>
    /// The value at a row-major index within a target <paramref name="rows"/>×<paramref name="columns"/>
    /// shape: a scalar broadcasts to every position; an array of the SAME shape yields its element; an
    /// array of a DIFFERENT shape is a dimension mismatch and yields <c>#VALUE!</c> (Excel parity).
    /// </summary>
    public abstract ComputedValue At(int index, int rows, int columns);
}

internal sealed class ScalarOperand : ArrayOperand
{
    private readonly ComputedValue _value;

    public ScalarOperand(ComputedValue value) => _value = value;

    public override bool IsArray => false;
    public override int Rows => 0;
    public override int Columns => 0;
    public override ComputedValue Scalar => _value;

    public override ComputedValue At(int index, int rows, int columns) => _value;
}

// A closed range: its origin and the sheet handle are resolved ONCE (like the previous ExpandRange), then
// each element reads numerically through the dense accessor — identical per-cell memoization, cycle guard
// and volatile taint. Row-major: index → (row, column) → (originRow + row, originColumn + column).
internal sealed class RangeOperand : ArrayOperand
{
    private readonly Workbook _workbook;
    private readonly int _handle;
    private readonly string _sheetName;
    private readonly int _originColumn;
    private readonly int _originRow;
    private readonly int _rows;
    private readonly int _columns;

    public RangeOperand(
        Workbook workbook,
        int handle,
        string sheetName,
        int originColumn,
        int originRow,
        int rows,
        int columns
    )
    {
        _workbook = workbook;
        _handle = handle;
        _sheetName = sheetName;
        _originColumn = originColumn;
        _originRow = originRow;
        _rows = rows;
        _columns = columns;
    }

    public override bool IsArray => true;
    public override int Rows => _rows;
    public override int Columns => _columns;

    public override ComputedValue At(int index, int rows, int columns)
    {
        if (_rows != rows || _columns != columns)
        {
            return ComputedValue.Error(Error.Value);
        }

        var row = index / _columns;
        var column = index % _columns;

        return _workbook.GetCellValueDense(
            _handle,
            _sheetName,
            _originColumn + column,
            _originRow + row
        );
    }
}

// Which axis of a rectangle a PositionNumbersOperand reports.
internal enum PositionAxis
{
    Row,
    Column,
}

// ROW(range)/COLUMN(range): every cell in row r shares the same worksheet row number (TopRow + r), and
// every cell in column c the same column number (LeftColumn + c). ONE class for both axes, because that
// is the whole difference between them: decomposing a row-major index into its rectangle coordinates
// gives the row as `index / _columns` and the column as `index % _columns`, so the operand needs the
// origin of its own axis and which of the two divisions to use. Splitting it in two bought nothing but a
// second copy of the mismatch guard.
internal sealed class PositionNumbersOperand : ArrayOperand
{
    private readonly int _origin;
    private readonly PositionAxis _axis;
    private readonly int _rows;
    private readonly int _columns;

    public PositionNumbersOperand(int origin, PositionAxis axis, int rows, int columns)
    {
        _origin = origin;
        _axis = axis;
        _rows = rows;
        _columns = columns;
    }

    public override bool IsArray => true;
    public override int Rows => _rows;
    public override int Columns => _columns;

    public override ComputedValue At(int index, int rows, int columns)
    {
        if (_rows != rows || _columns != columns)
        {
            return ComputedValue.Error(Error.Value);
        }

        return ComputedValue.Number(
            _origin + (_axis is PositionAxis.Row ? index / _columns : index % _columns)
        );
    }
}

internal sealed class BinaryOperand : ArrayOperand
{
    private readonly BinaryOperator _operator;
    private readonly ArrayOperand _left;
    private readonly ArrayOperand _right;
    private readonly int _rows;
    private readonly int _columns;

    public BinaryOperand(
        BinaryOperator @operator,
        ArrayOperand left,
        ArrayOperand right,
        int rows,
        int columns
    )
    {
        _operator = @operator;
        _left = left;
        _right = right;
        _rows = rows;
        _columns = columns;
    }

    public override bool IsArray => true;
    public override int Rows => _rows;
    public override int Columns => _columns;

    public override ComputedValue At(int index, int rows, int columns)
    {
        if (_rows != rows || _columns != columns)
        {
            return ComputedValue.Error(Error.Value);
        }

        return BinaryOperation.Apply(
            _operator,
            _left.At(index, _rows, _columns),
            _right.At(index, _rows, _columns)
        );
    }
}

internal sealed class IfOperand : ArrayOperand
{
    private readonly ArrayOperand _condition;
    private readonly ArrayOperand _whenTrue;
    private readonly ArrayOperand _whenFalse;
    private readonly int _rows;
    private readonly int _columns;

    public IfOperand(
        ArrayOperand condition,
        ArrayOperand whenTrue,
        ArrayOperand whenFalse,
        int rows,
        int columns
    )
    {
        _condition = condition;
        _whenTrue = whenTrue;
        _whenFalse = whenFalse;
        _rows = rows;
        _columns = columns;
    }

    public override bool IsArray => true;
    public override int Rows => _rows;
    public override int Columns => _columns;

    public override ComputedValue At(int index, int rows, int columns)
    {
        if (_rows != rows || _columns != columns)
        {
            return ComputedValue.Error(Error.Value);
        }

        var conditionValue = _condition.At(index, _rows, _columns);

        // A condition that is (or coerces from) an error propagates that error at this position.
        if (conditionValue.CoerceToBool(out var taken) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return taken ? _whenTrue.At(index, _rows, _columns) : _whenFalse.At(index, _rows, _columns);
    }
}
