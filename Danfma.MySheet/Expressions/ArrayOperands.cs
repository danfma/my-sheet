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
    /// extent: a scalar broadcasts to every position; an array is read through
    /// <see cref="Broadcasting.TryProject"/> — an axis of extent 1 repeats along the target's, and a
    /// position the array does not cover is <c>#N/A</c> (Excel's rule, measured on Aspose.Cells 26.6.0,
    /// 2026-09-10, CSE column — see <see cref="Broadcasting"/>).
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
        if (!Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own))
        {
            return ComputedValue.Error(Error.NA);
        }

        var row = own / _columns;
        var column = own % _columns;

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

// ROW(range)/COLUMN(range) over a rectangle: a VECTOR along the operand's own axis — ROW(A1:C3) is the 3x1
// column [1,2,3] and COLUMN(A1:C3) the 1x3 row [1,2,3], never an MxN rectangle. That is Excel's shape,
// measured on Aspose.Cells 26.6.0 (2026-09-10, CSE column): SUM(ROW(A1:C3)) = 6, COUNT(ROW(A1:C3)) = 3,
// SUM(COLUMN(A1:C3)) = 6. Phase 1 fabricated the rectangle ("every cell in row r shares the row number")
// only because the operand tree could not broadcast; with Broadcasting.TryProject the vector's other axis
// repeats, so the rectangle idioms still hold — SUM(ROW(A1:C3)*E1:E3) = 14, SUM(ROW(A1:A3)*COLUMN(A1:C1))
// = 36 (the outer product) — while the element count is the vector's. ONE class for both axes, because
// that is the whole difference between them: with the other axis of extent 1, `own / _columns` is the row
// of a column vector and `own % _columns` the column of a row vector, so the operand needs only its axis's
// origin and which of the two divisions to use. Splitting it in two bought nothing but a second copy of
// the projection.
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
        if (!Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own))
        {
            return ComputedValue.Error(Error.NA);
        }

        return ComputedValue.Number(
            _origin + (_axis is PositionAxis.Row ? own / _columns : own % _columns)
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

// ==============================================================================================
// Phase 8 — elementwise lifting: a Negate/Percent over an array, and a pure-scalar built-in over an array.
// ==============================================================================================

/// <summary>
/// The ONE mutable node in the engine — an argument slot for <see cref="LiftedFunctionOperand"/>. The lifted
/// built-in's node is created once per evaluation over these slots, and before each element is computed the
/// slots are rebound to that element's argument values, so the node re-evaluates per element without a fresh
/// node, a fresh argument array or a fresh literal per element (measured: 0 B/element, against 80-112 B for
/// a node rebuilt per element). It carries a <see cref="ComputedValue"/> rather than being one of the typed
/// literals (<see cref="NumberValue"/>, <see cref="StringValue"/>, …) because those are not TOTAL over the
/// value kinds: there is no literal node for <see cref="ComputedValueKind.Reference"/>, which an element can
/// be (a <c>+A1:A3</c> operand, a host custom function), while a value-carrying literal represents every kind.
/// </summary>
/// <remarks>
/// Mutability is acceptable here, and only here, because the instance never escapes the operand that created
/// it: it is reachable only from that operand's private array, is discarded with it, is never serialized (no
/// <c>[MemoryPackable]</c> attribute and NO <c>MemoryPackUnion</c> tag on <see cref="Expression"/>), and is
/// never hashed, compared or stored — the engine holds no <c>Dictionary&lt;Expression,…&gt;</c>,
/// <c>HashSet&lt;Expression&gt;</c> or <c>ConditionalWeakTable</c>, and the
/// <c>ArrayLifting.Elementwise</c> classification is the guarantee that a lifted node only ever calls
/// <see cref="Evaluate"/> on its argument nodes.
/// </remarks>
internal sealed record ScratchLiteral : ValueExpression
{
    public ComputedValue Value;

    public override ComputedValue Evaluate(EvaluationContext context) => Value;
}

// Negate/Percent over an array: BinaryOperand's shape guard, then the same coerce-then-apply ladder
// UnaryOperation.Evaluate runs for its non-Plus operators (UnaryOperation.Apply). Plus is NEVER built into
// this operand: it is Excel's reference-preserving no-op, routed through NamedReferences.CaptureValue so a
// range comes back as a REFERENCE value (SUM(+A1:A3) reads the cells today, and must keep doing so), and the
// builder's pattern excludes it so `+range` stays the opaque scalar that carries that reference.
internal sealed class UnaryOperand : ArrayOperand
{
    private readonly UnaryOperator _operator;
    private readonly ArrayOperand _inner;
    private readonly int _rows;
    private readonly int _columns;

    public UnaryOperand(UnaryOperator @operator, ArrayOperand inner, int rows, int columns)
    {
        _operator = @operator;
        _inner = inner;
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

        return UnaryOperation.Apply(_operator, _inner.At(index, _rows, _columns));
    }
}

// A pure-scalar built-in (an ArrayLifting.Elementwise registry entry) over at least one array argument. The
// node is created ONCE, through the registry's own factory, over a slot per argument; At() rebinds the slots
// to the element's values and evaluates that one node — the scalar body of the function is reused verbatim,
// so the lifted answer is the scalar answer element by element. The shape rule is BinaryOperand's, applied
// N-ary by the builder: scalars broadcast; arrays must share the shape, or the mismatched ones answer the
// #VALUE! marker from their own At() guard, which the body then receives as a VALUE — an error-propagating
// body (LEN, ROUND, LEFT, arithmetic) fills the result with #VALUE!, while an error-consuming body (IFERROR,
// IS*, N, T, IFS, SWITCH) sees the marker as its error argument and keeps going (today's behaviour, pinned,
// ahead of the broadcasting phase). The guard below is this operand's OWN shape check for its consumer.
//
// An OMITTED optional argument is the one slot that is not scratch. The parser leaves a literal BlankValue in
// it, and ten of the lifted built-ins (FIXED, DOLLAR, NUMBERVALUE, TEXTBEFORE/TEXTAFTER, VALUETOTEXT, the
// REGEX family, ADDRESS) detect the omission by pattern-matching that node — `Arguments[1] is not
// BlankValue` — so a scratch slot there would silently turn "omitted" into "blank, coerced to 0" (measured:
// FIXED(A1:A3,,TRUE) lost its decimals). The original node is handed through instead; being a literal it is
// never per-element, so nothing is lost.
internal sealed class LiftedFunctionOperand : ArrayOperand
{
    private readonly Expression _node;
    private readonly ScratchLiteral?[] _scratch;
    private readonly ArrayOperand[] _arguments;
    private readonly EvaluationContext _context;
    private readonly int _rows;
    private readonly int _columns;

    public LiftedFunctionOperand(
        Func<Expression[], Expression> create,
        Expression[] arguments,
        ArrayOperand[] operands,
        EvaluationContext context,
        int rows,
        int columns
    )
    {
        _arguments = operands;
        _context = context;
        _rows = rows;
        _columns = columns;
        _scratch = new ScratchLiteral?[operands.Length];

        var slots = new Expression[operands.Length];

        for (var j = 0; j < slots.Length; j++)
        {
            slots[j] =
                arguments[j] is BlankValue ? arguments[j] : _scratch[j] = new ScratchLiteral();
        }

        _node = create(slots);
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

        for (var j = 0; j < _scratch.Length; j++)
        {
            if (_scratch[j] is { } slot)
            {
                slot.Value = _arguments[j].At(index, _rows, _columns);
            }
        }

        return _node.Evaluate(_context);
    }
}
