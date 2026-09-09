using Danfma.MySheet.Expressions.Logical;
using Danfma.MySheet.Expressions.Lookup;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// The rectangular result of evaluating a sub-expression element-by-element: a vector of
/// <see cref="ComputedValue"/> laid out ROW-MAJOR with its <see cref="Rows"/>×<see cref="Columns"/>
/// dimensions. Errors and logicals are preserved per element — the CONSUMER decides how to fold/propagate
/// them (Phase B); this struct never collapses the vector to a scalar.
/// </summary>
internal readonly struct ArrayEvaluationResult
{
    public int Rows { get; }
    public int Columns { get; }
    public ComputedValue[] Values { get; }

    public ArrayEvaluationResult(int rows, int columns, ComputedValue[] values)
    {
        Rows = rows;
        Columns = columns;
        Values = values;
    }

    /// <summary>The element count (<c>Rows * Columns</c>, i.e. <c>Values.Length</c>).</summary>
    public int Length => Values.Length;
}

/// <summary>
/// The internal "mini-CSE" element-wise evaluator (Phase A of <c>plans/mini-cse-array-arguments.md</c>).
/// Given an AST node and a context it reproduces Excel's implicit array/CSE semantics WITHOUT any public
/// array value, spilling or new AST node: a closed range becomes a vector of its per-cell values, scalars
/// broadcast, <c>BinaryOperation</c> zips element-wise, <c>IF</c> zips a condition array against its
/// branches (a branch-less <c>IF</c> yields a logical <c>FALSE</c> where the condition is false — the idiom
/// <c>SMALL(IF(...))</c> depends on this), and <c>ROW</c>/<c>COLUMN</c> of a rectangle becomes a vector of
/// row/column numbers — over a range written literally, and over a defined name or any other node that
/// DENOTES one (a structured reference, a <c>':'</c> range with reference-returning endpoints), whose shape
/// is discovered by resolving it. Any node outside this set is treated as a scalar (broadcast); a
/// whole-column/open range is REFUSED (the cost guard) so the whole evaluation reports "not an array" and
/// the caller keeps its current scalar path.
///
/// <para>Two consumption shapes share ONE recursive builder (<see cref="TryBuildOperand"/>): the LAZY
/// <see cref="ArrayStream"/> (element-on-demand, no vector — used by the aggregating consumers SUM/COUNT/
/// AVERAGE/MIN/MAX, SMALL/LARGE and INDEX so a 50k-row idiom allocates only the handful of tree nodes, not a
/// <c>ComputedValue[50k]</c>), and the eager <see cref="ArrayEvaluationResult"/> (materialized once from the
/// same tree) that the direct unit tests drive. Because both walk the identical operand tree, the streamed
/// element sequence is bit-for-bit the materialized vector, in the SAME row-major order — so error
/// propagation ("first error, scan order"), broadcast, dimension-mismatch <c>#VALUE!</c> and the
/// evaluate-scalar-operands-once rule (a volatile branch draws once and broadcasts) are preserved.</para>
/// </summary>
internal static class ArrayEvaluation
{
    /// <summary>
    /// Tries to evaluate <paramref name="expression"/> as an array, MATERIALIZING it into a row-major vector.
    /// Returns <c>true</c> (with the vector in <paramref name="result"/>) only when the sub-tree genuinely
    /// produces an array — a closed range, a <c>ROW</c>/<c>COLUMN</c> of one, or an operation/IF with at least
    /// one array operand. A bare scalar, a non-eligible node, or anything touching an open range returns
    /// <c>false</c> (the current scalar path is untouched). Production consumers use
    /// <see cref="TryEvaluateStream"/>; this
    /// eager form exists for the direct unit tests of the vector semantics.
    /// </summary>
    public static bool TryEvaluate(
        Expression expression,
        EvaluationContext context,
        out ArrayEvaluationResult result
    )
    {
        if (TryBuildOperand(expression, context, out var operand) && operand.IsArray)
        {
            var rows = operand.Rows;
            var columns = operand.Columns;
            var values = new ComputedValue[rows * columns];

            for (var index = 0; index < values.Length; index++)
            {
                values[index] = operand.At(index, rows, columns);
            }

            result = new ArrayEvaluationResult(rows, columns, values);
            return true;
        }

        result = default;
        return false;
    }

    /// <summary>
    /// Tries to build a LAZY element-wise view of <paramref name="expression"/> — the allocation-free twin of
    /// <see cref="TryEvaluate"/>. Succeeds under the exact same conditions, but instead of a materialized
    /// vector it returns an <see cref="ArrayStream"/> that computes each element on demand from a small
    /// operand tree (resolved once): the aggregating consumers enumerate it without ever allocating a
    /// <c>ComputedValue[]</c>. Scalar sub-expressions are still evaluated ONCE at build time (broadcast), so a
    /// volatile branch/operand draws a single value and its taint lands in the enclosing cell frame exactly as
    /// before.
    /// </summary>
    public static bool TryEvaluateStream(
        Expression expression,
        EvaluationContext context,
        out ArrayStream stream
    )
    {
        if (TryBuildOperand(expression, context, out var operand) && operand.IsArray)
        {
            stream = new ArrayStream(operand, operand.Rows, operand.Columns);
            return true;
        }

        stream = default;
        return false;
    }

    /// <summary>
    /// A CHEAP pre-check — it never evaluates <paramref name="expression"/> itself — for whether it would
    /// produce an array through the eligible structural set. Consumers gate on this BEFORE calling
    /// <see cref="TryEvaluate"/>/<see cref="TryEvaluateStream"/>, so the scalar hot path pays only a shallow
    /// type-walk and never a double evaluation: when this returns <c>false</c> the consumer keeps its existing
    /// scalar path untouched; when it returns <c>true</c> the subsequent build is guaranteed to succeed and is
    /// the SINGLE evaluation of the argument. It mirrors <see cref="Probe"/>/<see cref="TryBuildOperand"/>
    /// exactly (same array-producing cases, same open-range refusal), so it is true iff the build succeeds as
    /// an array.
    /// </summary>
    /// <remarks>
    /// It takes a <paramref name="context"/> because the check is no longer purely SYNTACTIC: whether
    /// <c>ROW</c>/<c>COLUMN</c> of a defined name (or of any other node that merely DENOTES a reference) is an
    /// array or a scalar depends on the SHAPE the argument resolves to — a rectangle vs a single cell — which
    /// no type-walk can know. That case costs one reference RESOLUTION — a dictionary lookup plus a virtual
    /// call for a name, never an evaluation of the <c>ROW</c>/<c>COLUMN</c> node itself, though resolving a
    /// <c>':'</c> range with reference-returning endpoints does evaluate those endpoints' own arguments — and
    /// it is the only case that costs more than the type-walk. See <see cref="ResolvePositionRange"/>.
    /// </remarks>
    public static bool IsArrayEligible(Expression expression, EvaluationContext context) =>
        Probe(expression, context).IsArray;

    // The pure-shape twin of TryBuildOperand: decides, WITHOUT evaluating the expression, whether the build
    // would succeed (Succeeds — no refused open range on the eligible path) and whether the result is an array
    // (IsArray). Must track the builder's structure exactly so IsArrayEligible == (build result).
    //
    // One case cannot be answered from the syntax alone — ROW/COLUMN over a node that only DENOTES a
    // reference — so both this probe and the builder ask the SAME oracle, ResolvePositionRange, on the same
    // argument in the same context, and therefore agree. The oracle resolves the argument rather than
    // evaluating the ROW/COLUMN node, and the build resolves it a second time, which is what keeps a
    // reference-returning FUNCTION argument out of that arm: see ResolvePositionRange.
    private static (bool Succeeds, bool IsArray) Probe(
        Expression expression,
        EvaluationContext context
    )
    {
        switch (expression)
        {
            case RangeReference:
                return (true, true);

            // Phase 2 audit (shared-formula delta production): a range written INSIDE a shared-formula
            // master is an anchored node, not a plain RangeReference — without this arm it fell through to
            // the `default` case below (an opaque scalar, evaluated once via AnchoredRangeReference.Evaluate,
            // which always yields #VALUE! since a range has no scalar value), silently breaking the mini-CSE
            // idiom (e.g. SMALL(IF(A1:A3=…, …))) for a shared-formula slave.
            case AnchoredRangeReference:
                return (true, true);

            // An open/whole-column range in an array position is refused (the cost guard).
            case OpenRangeReference:
                return (false, false);

            // ROW(x)/COLUMN(x) over a reference: ONE arm each, because the two functions differ only in the
            // AXIS the shared shape walk reports — see ProbePosition. The `[NameReference or Reference]`
            // pattern admits every reference NODE (RangeReference, AnchoredRangeReference and
            // OpenRangeReference all derive from Reference) plus a defined name, and deliberately not a
            // reference-returning FUNCTION: see ResolvePositionRange.
            case Row { Arguments: [NameReference or Reference] } row:
                return ProbePosition(row.Arguments[0], context);

            case Column { Arguments: [NameReference or Reference] } column:
                return ProbePosition(column.Arguments[0], context);

            case BinaryOperation binary:
            {
                var left = Probe(binary.Left, context);
                if (!left.Succeeds)
                {
                    return (false, false);
                }

                var right = Probe(binary.Right, context);
                if (!right.Succeeds)
                {
                    return (false, false);
                }

                return (true, left.IsArray || right.IsArray);
            }

            case If ifNode when ifNode.Arguments.Length is 2 or 3:
            {
                var condition = Probe(ifNode.Arguments[0], context);
                if (!condition.Succeeds)
                {
                    return (false, false);
                }

                // A scalar condition makes the IF an opaque scalar (its native short-circuit applies): it
                // succeeds but is not an array, so the branches are never probed.
                if (!condition.IsArray)
                {
                    return (true, false);
                }

                if (!Probe(ifNode.Arguments[1], context).Succeeds)
                {
                    return (false, false);
                }

                if (ifNode.Arguments.Length == 3 && !Probe(ifNode.Arguments[2], context).Succeeds)
                {
                    return (false, false);
                }

                return (true, true);
            }

            // Anything else is an opaque scalar: succeeds (evaluated once when actually built), not an array.
            default:
                return (true, false);
        }
    }

    // The shape twin of TryBuildPositionOperand — same argument order, same three outcomes, so ROW/COLUMN
    // stay eligible exactly when their build succeeds as an array. A rectangle (written literally, or written
    // inside a shared-formula master as an anchored node) is an array with no resolution at all; an open range
    // is the cost-guard refusal; anything else asks the oracle, whose Array/Refused/Scalar map onto the same
    // three answers the builder gives (a Scalar being the opaque-scalar answer: it succeeds, broadcast).
    private static (bool Succeeds, bool IsArray) ProbePosition(
        Expression argument,
        EvaluationContext context
    ) =>
        argument switch
        {
            RangeReference or AnchoredRangeReference => (true, true),
            OpenRangeReference => (false, false),
            _ => ResolvePositionRange(argument, context, out _) switch
            {
                PositionArgumentShape.Array => (true, true),
                PositionArgumentShape.Refused => (false, false),
                _ => (true, false),
            },
        };

    // ==============================================================================================
    // The lazy operand tree: built once (array leaves resolved once, scalar sub-expressions evaluated once),
    // then each element is computed on demand. The classes mirror the previous materialized `Operand.At`
    // element-for-element, so both consumption shapes agree bit-for-bit.
    // ==============================================================================================

    /// <summary>
    /// A recursively-built operand: either a scalar (to broadcast) or a rectangular array whose elements are
    /// computed on demand by <see cref="At"/>. Building FAILS (returns <c>false</c> from
    /// <see cref="TryBuildOperand"/>) only when the sub-tree reaches an open/whole-column range in an array
    /// position — the cost guard — so the whole evaluation degrades to "not an array".
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

    private sealed class ScalarOperand : ArrayOperand
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
    private sealed class RangeOperand : ArrayOperand
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
    private enum PositionAxis
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
    private sealed class PositionNumbersOperand : ArrayOperand
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

    private sealed class BinaryOperand : ArrayOperand
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

    private sealed class IfOperand : ArrayOperand
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

            return taken
                ? _whenTrue.At(index, _rows, _columns)
                : _whenFalse.At(index, _rows, _columns);
        }
    }

    private static bool TryBuildOperand(
        Expression expression,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        switch (expression)
        {
            case RangeReference range:
                operand = BuildRange(range, context);
                return true;

            // Phase 2 audit (shared-formula delta production): mirrors the RangeReference case above for a
            // range written INSIDE a shared-formula master (an anchored node) — resolved to its concrete,
            // delta-applied twin so BuildRange sees an ordinary RangeReference, no logic duplicated.
            case AnchoredRangeReference anchoredRange:
                operand = BuildRange(anchoredRange.ToRangeReference(context), context);
                return true;

            // Whole-column / whole-row / one-sided open reference: the cost guard keeps it OUT of the
            // mini-CSE. Refuse so the whole evaluation reports "not an array".
            case OpenRangeReference:
                operand = null!;
                return false;

            // ROW(x)/COLUMN(x) over a reference: ONE arm each, both walking the same shapes on their own
            // axis — see TryBuildPositionOperand, whose order Probe/ProbePosition mirrors exactly. The
            // pattern admits every reference NODE plus a defined name, and deliberately not a
            // reference-returning FUNCTION: see ResolvePositionRange.
            case Row { Arguments: [NameReference or Reference] } row:
                return TryBuildPositionOperand(
                    row,
                    row.Arguments[0],
                    PositionAxis.Row,
                    context,
                    out operand
                );

            case Column { Arguments: [NameReference or Reference] } column:
                return TryBuildPositionOperand(
                    column,
                    column.Arguments[0],
                    PositionAxis.Column,
                    context,
                    out operand
                );

            case BinaryOperation binary:
                return TryBuildBinary(binary, context, out operand);

            case If ifNode when ifNode.Arguments.Length is 2 or 3:
                return TryBuildIf(ifNode, context, out operand);

            // Anything else is an opaque scalar: evaluate ONCE and broadcast. (This is where nested scalar
            // functions — SUM(A:A), a bare cell, a literal, a volatile RAND() — enter, without recursing.)
            default:
                operand = new ScalarOperand(expression.Evaluate(context));
                return true;
        }
    }

    // The build for ROW(x)/COLUMN(x) on one axis, walking the shapes in the SAME order ProbePosition does so
    // the two cannot drift: a rectangle written literally; one written INSIDE a shared-formula master (an
    // anchored node, resolved to its concrete delta-applied twin); the open-range refusal (the cost guard);
    // then anything that merely DENOTES a reference, through the oracle.
    //
    // The two rectangle arms are FAST PATHS, not requirements: they exist so the common case pays no
    // resolution at all. The anchored one was a correctness fix when it was written (a slave's ROW(A1:A3) fell
    // to the opaque-scalar branch, and AnchoredRangeReference.Evaluate is always #VALUE! since a range has no
    // scalar value), but the oracle arm below now resolves an anchored range to the same delta-applied
    // rectangle — verified by deleting both anchored arms and watching SharedFormulaSlaveFunctionTests stay
    // green. So no assertion can distinguish them from the fallback; what the tests there pin is the delta.
    //
    // `node` is the ROW/COLUMN call and `argument` its single argument: Function declares no Arguments member,
    // so the caller — which has already destructured the argument to pattern-match it — passes both. The node
    // is needed only on the Scalar path, to evaluate itself once.
    //
    // The `default` arm below therefore only ever sees the caller's narrowed set (a defined name, a cell, a
    // union, a DynamicRange), NEVER a reference-returning function; widening either caller's pattern would
    // change that, which ResolvePositionRange explains is not free.
    private static bool TryBuildPositionOperand(
        Expression node,
        Expression argument,
        PositionAxis axis,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        switch (argument)
        {
            case RangeReference range:
                operand = PositionOperand(range.GetBounds(), axis);
                return true;

            case AnchoredRangeReference anchored:
                operand = PositionOperand(anchored.ToRangeReference(context).GetBounds(), axis);
                return true;

            case OpenRangeReference:
                operand = null!;
                return false;

            default:
            {
                var shape = ResolvePositionRange(argument, context, out var bounds);

                if (shape is PositionArgumentShape.Refused)
                {
                    operand = null!;
                    return false;
                }

                operand =
                    shape is PositionArgumentShape.Array
                        ? PositionOperand(bounds, axis)
                        // A single cell, a union, or a name that resolves to no reference at all: let the
                        // node evaluate itself ONCE and broadcast, which is exactly its (correct) scalar
                        // answer, error included.
                        : new ScalarOperand(node.Evaluate(context));
                return true;
            }
        }
    }

    // The operand over one axis of a resolved rectangle: that axis's origin (TopRow or LeftColumn) is what
    // every cell along the other axis shares.
    private static ArrayOperand PositionOperand(RangeBounds bounds, PositionAxis axis) =>
        new PositionNumbersOperand(
            axis is PositionAxis.Row ? bounds.TopRow : bounds.LeftColumn,
            axis,
            bounds.RowCount,
            bounds.ColumnCount
        );

    // What a ROW/COLUMN argument that DENOTES a reference contributes to the mini-CSE, once resolved.
    private enum PositionArgumentShape
    {
        // A rectangle: ROW/COLUMN over it is a vector of positions, one per cell.
        Array,

        // No rectangle to spread over — a single cell, a union, a name that does not resolve: the ROW/COLUMN
        // node evaluates itself once and broadcasts its own scalar answer (or its own error).
        Scalar,

        // An open/whole-column reference: the cost guard refuses the whole array evaluation.
        Refused,
    }

    // Resolves a ROW/COLUMN argument to the rectangle whose positions the operand reports. This is the ONE
    // oracle Probe and TryBuildOperand share: the shape of such an argument is context-dependent, so a
    // syntactic probe could not answer it, and a probe that GUESSED (leaving the build to fail) would break
    // the documented "eligible iff the build succeeds — and the build is the SINGLE evaluation" contract.
    //
    // Because BOTH callers resolve, every argument admitted here is resolved TWICE per evaluation. For the
    // shapes the arms admit that is a scope/dictionary lookup and a virtual call — idempotent and cheap. It is
    // exactly why a reference-returning FUNCTION argument (ROW(INDEX(…))/OFFSET/INDIRECT) is NOT admitted:
    // resolving one evaluates the function's own arguments, so the pair of resolutions would draw a volatile
    // twice where the scalar path draws it once (measured: a counting function inside ROW(INDEX(…,TICK(),1))
    // is called exactly once, before and after this fix). Such an argument stays an opaque scalar here — still
    // its correct scalar answer (the resolved reference's top row / leftmost column), just not yet its array
    // shape.
    //
    // The one admitted shape whose resolution is not free is a DynamicRange with reference-returning ENDPOINTS
    // (ROW(INDEX(A1:A3,1,1):A3)), in scope deliberately as the same `or Reference` hook a structured table
    // column will land on. It costs no EXTRA draw: the scalar path this replaces already resolved it twice
    // (ReferenceGuard.MissingSheet, then ReferencePosition), so the count is 2 either way — measured. What a
    // SIDE-EFFECTING endpoint does expose is that the two resolutions can then disagree on the rectangle, and
    // the build's answer wins (ROW(INDEX(A1:A3,TICK(),1):A3) folds to 5, where Excel's single evaluation gives
    // 6 — and where the pre-fix scalar answer was a likewise-wrong 2). A pure endpoint, which is every real
    // one, resolves to the same rectangle twice and is exact.
    //
    // boundOpenRanges:false keeps an open range OPEN so it reaches the Refused arm instead of being silently
    // collapsed to its populated bounding box — both the cost guard this class states and the same choice
    // ROW/COLUMN's own scalar fallback (ReferencePosition) makes, where ROW(A:A) is the DECLARED row 1.
    private static PositionArgumentShape ResolvePositionRange(
        Expression argument,
        EvaluationContext context,
        out RangeBounds bounds
    )
    {
        bounds = default;

        if (
            !NamedReferences.TryResolveReference(
                argument,
                context,
                out var reference,
                boundOpenRanges: false
            )
        )
        {
            // Scalar, not Refused: the enclosing ROW/COLUMN then evaluates once and reports the ARGUMENT's own
            // error (#NAME? for an unknown name) through ReferencePosition. Refusing instead would hand the
            // consumer's scalar path an operand tree it never built, losing that error in an enclosing
            // operation (SUM(A1:A3+ROW(Nope)) must be #NAME?, not the #VALUE! of a range in a scalar add).
            return PositionArgumentShape.Scalar;
        }

        // A reference that resolves onto a DELETED sheet is a structural #REF!, which ROW/COLUMN's own
        // ReferenceGuard pass reports — degrade to Scalar so that error is what gets broadcast, rather than
        // handing back bounds and a plausible row number for a sheet that no longer exists. (The syntactic
        // arms above cannot reach this: a literal Ghost!A1:A3 carries its dead sheet name into the operand,
        // a pre-existing gap this fix neither widens nor closes.)
        if (ReferenceGuard.MissingSheet(reference, context) is not null)
        {
            return PositionArgumentShape.Scalar;
        }

        switch (reference)
        {
            case RangeReference range:
                // GetBounds() parses BOTH corners, so it is called once here and the rectangle handed back.
                bounds = range.GetBounds();
                return PositionArgumentShape.Array;

            case OpenRangeReference:
                return PositionArgumentShape.Refused;

            // A single cell (1x1 — nothing to spread over) or a union (no single position: #VALUE!).
            default:
                return PositionArgumentShape.Scalar;
        }
    }

    private static ArrayOperand BuildRange(RangeReference range, EvaluationContext context)
    {
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(range.SheetName);
        var bounds = range.GetBounds();

        return new RangeOperand(
            workbook,
            handle,
            range.SheetName,
            bounds.LeftColumn,
            bounds.TopRow,
            bounds.RowCount,
            bounds.ColumnCount
        );
    }

    private static bool TryBuildBinary(
        BinaryOperation binary,
        EvaluationContext context,
        out ArrayOperand operand
    )
    {
        if (
            !TryBuildOperand(binary.Left, context, out var left)
            || !TryBuildOperand(binary.Right, context, out var right)
        )
        {
            operand = null!;
            return false;
        }

        // Neither side is an array → the whole operation is a scalar (broadcast), evaluated once.
        if (!left.IsArray && !right.IsArray)
        {
            operand = new ScalarOperand(
                BinaryOperation.Apply(binary.Operator, left.Scalar, right.Scalar)
            );
            return true;
        }

        var (rows, columns) = ResultShape(left, right);
        operand = new BinaryOperand(binary.Operator, left, right, rows, columns);
        return true;
    }

    private static bool TryBuildIf(If ifNode, EvaluationContext context, out ArrayOperand operand)
    {
        if (!TryBuildOperand(ifNode.Arguments[0], context, out var condition))
        {
            operand = null!;
            return false;
        }

        // A scalar condition is not an array selection: treat the whole IF as an opaque scalar (its native
        // short-circuit Evaluate applies, evaluated once). Only an ARRAY condition drives the element-wise zip.
        if (!condition.IsArray)
        {
            operand = new ScalarOperand(ifNode.Evaluate(context));
            return true;
        }

        if (!TryBuildOperand(ifNode.Arguments[1], context, out var whenTrue))
        {
            operand = null!;
            return false;
        }

        // IF without an else branch yields a logical FALSE where the condition is false (Excel), which the
        // aggregators then ignore — the whole point of the SMALL(IF(...)) idiom.
        var hasElse = ifNode.Arguments.Length == 3;
        ArrayOperand whenFalse = new ScalarOperand(ComputedValue.Boolean(false));
        if (hasElse && !TryBuildOperand(ifNode.Arguments[2], context, out whenFalse))
        {
            operand = null!;
            return false;
        }

        operand = new IfOperand(condition, whenTrue, whenFalse, condition.Rows, condition.Columns);
        return true;
    }

    // The shape of a binary result: a scalar takes the other side's shape; two equal-shaped arrays keep it;
    // mismatched arrays produce the per-axis maximum, filled entirely with #VALUE! by the At() mismatch rule.
    private static (int Rows, int Columns) ResultShape(ArrayOperand left, ArrayOperand right)
    {
        if (!left.IsArray)
        {
            return (right.Rows, right.Columns);
        }

        if (!right.IsArray)
        {
            return (left.Rows, left.Columns);
        }

        return left.Rows == right.Rows && left.Columns == right.Columns
            ? (left.Rows, left.Columns)
            : (Math.Max(left.Rows, right.Rows), Math.Max(left.Columns, right.Columns));
    }

    /// <summary>
    /// A no-alloc, lazy view over the mini-CSE array: the resolved operand tree plus its row-major shape.
    /// <see cref="ElementAt"/> computes a single element on demand (INDEX walks to the n-th only); a
    /// <c>foreach</c> binds the struct <see cref="Enumerator"/> by duck typing (the <see cref="List{T}"/>
    /// pattern), so the aggregating consumers (SUM/COUNT/AVERAGE/MIN/MAX, SMALL/LARGE) scan every element
    /// row-major with no interface dispatch and no <c>ComputedValue[]</c>.
    /// </summary>
    internal readonly struct ArrayStream
    {
        private readonly ArrayOperand _root;

        internal ArrayStream(ArrayOperand root, int rows, int columns)
        {
            _root = root;
            Rows = rows;
            Columns = columns;
        }

        public int Rows { get; }
        public int Columns { get; }
        public int Length => Rows * Columns;

        /// <summary>The value at a row-major index (0-based), computed on demand.</summary>
        public ComputedValue ElementAt(int index) => _root.At(index, Rows, Columns);

        public Enumerator GetEnumerator() => new(_root, Rows, Columns);

        public struct Enumerator
        {
            private readonly ArrayOperand _root;
            private readonly int _rows;
            private readonly int _columns;
            private readonly int _length;
            private int _index;
            private ComputedValue _current;

            internal Enumerator(ArrayOperand root, int rows, int columns)
            {
                _root = root;
                _rows = rows;
                _columns = columns;
                _length = rows * columns;
                _index = -1;
                _current = default;
            }

            public readonly ComputedValue Current => _current;

            public bool MoveNext()
            {
                var next = _index + 1;
                if (next >= _length)
                {
                    return false;
                }

                _index = next;
                _current = _root.At(next, _rows, _columns);
                return true;
            }
        }
    }
}
