namespace Danfma.MySheet.Expressions;

// ==============================================================================================
// The lazy operand tree: built once (array leaves resolved once, scalar sub-expressions evaluated once),
// then each element is computed on demand. The classes mirror the previous materialized `Operand.At`
// element-for-element, so both consumption shapes agree bit-for-bit.
// ==============================================================================================

/// <summary>
/// The ONE extension point for a node that PRODUCES an array inside the mini-CSE — Phase 7's
/// <c>FILTER</c>, <c>SORT</c>, <c>UNIQUE</c> and <c>SEQUENCE</c>, and every later producer
/// (<c>TRANSPOSE</c>, <c>SORTBY</c>, <c>TAKE</c>, <c>HSTACK</c>, …). A node that implements it is reached by
/// exactly two arms — one in <see cref="ArrayEvaluation.Probe"/>, one in
/// <see cref="ArrayEvaluation.TryBuildOperand"/> — so adding a producer never edits
/// <c>ArrayEvaluation.cs</c>; it mirrors the codebase's extension-by-shared-abstraction style
/// (<c>INumericFold</c>). A producer's node keeps <see cref="Expression.Evaluate"/> as its CELL answer,
/// which is <see cref="ArrayEvaluation.FirstElement(Expression, EvaluationContext)"/> — Excel's <c>@</c> on an array, the top-left.
/// </summary>
/// <remarks>
/// <para>The two members are twins and MUST agree, because the mini-CSE's documented contract is
/// <c>IsArrayEligible == (the build succeeds as an array)</c> and <c>NumericAggregation.Fold</c> relies on
/// it to evaluate a volatile argument once: <see cref="ProbeArray"/> answers from SHAPE alone — it never
/// evaluates the node, though like <see cref="ArrayEvaluation.Probe"/> it may resolve a name — and
/// <see cref="TryBuildArrayOperand"/> returns <c>false</c> ONLY where the probe answered
/// <c>Succeeds = false</c>. Both recurse into their children through the widened
/// <see cref="ArrayEvaluation.Probe"/>/<see cref="ArrayEvaluation.TryBuildOperand"/> rather than by
/// pattern-matching the child node, so a name, a table column or another producer in the child slot is
/// whatever those two say it is.</para>
///
/// <para>A REFUSED child (an open range somewhere below — the cost guard) makes the producer refuse:
/// <c>(false, false)</c> from the probe, <c>false</c> from the build. That is the <c>BinaryOperation</c>/
/// <c>If</c> side of Phase 8's split, not the lift's opaque-scalar side, so the consumer keeps its scalar
/// path and reaches <see cref="ArrayEvaluation.FirstElement(Expression, EvaluationContext)"/>, which answers <c>#VALUE!</c>
/// (<c>ArrayProducerContractTests</c>; <c>SUM(FILTER(A:A,A:A>0))</c> is the phase's pinned deviation).</para>
///
/// <para>Whatever the build hands back when it succeeds is an ARRAY with <c>Rows >= 1 &amp;&amp; Columns >= 1</c>
/// — never a scalar, never a 0-extent shape: a 1x1 source, a bad argument and an EMPTY result are all a
/// 1x1 <see cref="SingletonArrayOperand"/> (see <see cref="ArrayShaping"/> for the invariant and what
/// breaks without it). The operand projects through <see cref="Broadcasting.TryProject"/> like every other
/// array operand, so a producer composes under an operator, a lifted function, an <c>IF</c> or another
/// producer with no code of its own.</para>
///
/// <para>Two mechanical rules for an implementation. (1) The interface and <see cref="ArrayOperand"/> are
/// internal while the function records are public, so a record MUST implement both members EXPLICITLY
/// (<c>bool IArrayProducer.TryBuildArrayOperand(…)</c>) — an implicit public method mentioning
/// <see cref="ArrayOperand"/> is a CS0050 inconsistent-accessibility error. (2) The registry entry MUST be
/// <see cref="ArrayLifting.Consumes"/>: the <c>Function when TryGetLift</c> arm sits BEFORE the producer
/// arm in both switches, so an <see cref="ArrayLifting.Elementwise"/> producer does not merely answer from
/// one cell — the lift asks for the node's scalar, <c>FirstElement</c> builds the array operand, that build
/// re-enters <c>TryGetLift</c>, and the cycle recurses until the STACK OVERFLOWS, taking the host down with
/// no failing test. Measured by this phase's final review. That is why
/// <c>FunctionRegistry.RequireProducerIsConsumes</c> now refuses the combination at the registry's static
/// initialization, naming the offender, and why an earlier version of this sentence understated it.</para>
/// </remarks>
internal interface IArrayProducer
{
    /// <summary>
    /// The shape twin of <see cref="TryBuildArrayOperand"/>: whether the build would succeed
    /// (<c>Succeeds</c> — no refused open range on the eligible path) and whether it yields an array
    /// (<c>IsArray</c> — always <c>true</c> when it succeeds, since a producer's result is never a scalar).
    /// Never evaluates the node.
    /// </summary>
    (bool Succeeds, bool IsArray) ProbeArray(EvaluationContext context);

    /// <summary>
    /// Builds the producer's operand — its child operands through
    /// <see cref="ArrayEvaluation.TryBuildOperand"/>, its scalar arguments evaluated ONCE here (the
    /// build-time read the laziness contract sanctions) — or returns <c>false</c> when a child build is
    /// refused. See the type remarks for what the operand must satisfy.
    /// </summary>
    bool TryBuildArrayOperand(EvaluationContext context, out ArrayOperand operand);
}

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
    /// extent: a scalar broadcasts to every position; an array operand is read through
    /// <see cref="Broadcasting.TryProject"/> — an axis of extent 1 repeats along the target's, and a
    /// position the array does not cover is <c>#N/A</c> (Excel's rule, measured on Aspose.Cells 26.6.0,
    /// 2026-09-10, CSE column — see <see cref="Broadcasting"/>). A COMPOSITE (<c>BinaryOperand</c>,
    /// <c>IfOperand</c>, <c>UnaryOperand</c>, <c>LiftedFunctionOperand</c>) projects first and then asks its
    /// children at its OWN index and extent, so nested vectors compose level by level:
    /// <c>SUM((A1:A3*H1:H2)*E5:G5)</c> is <c>#N/A</c> with six countable elements
    /// (<c>VectorBroadcastingTests</c>).
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
        if (!Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own))
        {
            return ComputedValue.Error(Error.NA);
        }

        return BinaryOperation.Apply(
            _operator,
            _left.At(own, _rows, _columns),
            _right.At(own, _rows, _columns)
        );
    }
}

// IF over an array condition: the extent is the fold of the condition AND both branches (the builder's
// ShapeFold), so IF(E1:E3>1,E5:G5,0) is 3x3 and IF(H1:H2>1,A1:C3,0) marks row 3 #N/A instead of quietly
// counting the else-values of a 2x1 extent. Each element reads the condition at the IF's own index, then
// ONLY the taken branch — the other branch is never read, which is what keeps a volatile branch drawn once.
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
        if (!Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own))
        {
            return ComputedValue.Error(Error.NA);
        }

        var conditionValue = _condition.At(own, _rows, _columns);

        // A condition that is (or coerces from) an error propagates that error at this position — an
        // uncovered condition included, so its #N/A wins over whatever the branches hold there. The text
        // "TRUE"/"FALSE" coerces here exactly as it does in If.Evaluate, so the array path of a condition
        // never disagrees with the scalar one.
        if (conditionValue.CoerceToBoolAllowingTextWords(out var taken) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return taken ? _whenTrue.At(own, _rows, _columns) : _whenFalse.At(own, _rows, _columns);
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

// Negate/Percent over an array: the same Broadcasting.TryProject projection as every array operand (#N/A
// where the operand does not cover the position, the inner operand asked at THIS operand's own index and
// extent), then the same coerce-then-apply ladder UnaryOperation.Evaluate runs for its non-Plus operators
// (UnaryOperation.Apply). Plus is NEVER built into
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
        if (!Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own))
        {
            return ComputedValue.Error(Error.NA);
        }

        return UnaryOperation.Apply(_operator, _inner.At(own, _rows, _columns));
    }
}

// A pure-scalar built-in (an ArrayLifting.Elementwise registry entry) over at least one array argument. The
// node is created ONCE, through the registry's own factory, over a slot per argument; At() rebinds the slots
// to the element's values and evaluates that one node — the scalar body of the function is reused verbatim,
// so the lifted answer is the scalar answer element by element. The shape rule is BinaryOperand's, applied
// N-ary by the builder's ShapeFold through Broadcasting.Axis: scalars broadcast, and every array argument is
// read through Broadcasting.TryProject at this operand's OWN index and extent, so an axis of extent 1
// repeats (SUM(ROUND(E1:E3,E5:G5)) = 18, the outer product) and a position an argument does not cover
// hands the body a real per-element #N/A — an error-propagating body (LEN, ROUND, LEFT, arithmetic) carries
// it to that element, while an error-consuming body (IFERROR, IS*, N, T, IFS, SWITCH) recovers it:
// SUM(IFERROR(A1:C3*H1:H2,0)) = 36, SUM(ISNA(A1:C3*H1:H2)*1) = 3 (VectorBroadcastingTests; Aspose.Cells
// 26.6.0, 2026-09-10, CSE column). The projection below is this operand's OWN, for its consumer's extent.
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
        if (!Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own))
        {
            return ComputedValue.Error(Error.NA);
        }

        for (var j = 0; j < _scratch.Length; j++)
        {
            if (_scratch[j] is { } slot)
            {
                slot.Value = _arguments[j].At(own, _rows, _columns);
            }
        }

        return _node.Evaluate(_context);
    }
}
