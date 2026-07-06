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
/// <c>SMALL(IF(...))</c> depends on this), and <c>ROW(range)</c> becomes a vector of row numbers. Any node
/// outside this set is treated as a scalar (broadcast); a whole-column/open range is REFUSED (the cost
/// guard) so the whole evaluation reports "not an array" and the caller keeps its current scalar path.
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
    /// produces an array — a closed range, a <c>ROW(range)</c>, or an operation/IF with at least one array
    /// operand. A bare scalar, a non-eligible node, or anything touching an open range returns <c>false</c>
    /// (the current scalar path is untouched). Production consumers use <see cref="TryEvaluateStream"/>; this
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
    /// A CHEAP syntactic pre-check — no evaluation — for whether <paramref name="expression"/> would
    /// produce an array through the eligible structural set. Consumers gate on this BEFORE calling
    /// <see cref="TryEvaluate"/>/<see cref="TryEvaluateStream"/>, so the scalar hot path pays only a shallow
    /// type-walk and never a double evaluation: when this returns <c>false</c> the consumer keeps its existing
    /// scalar path untouched; when it returns <c>true</c> the subsequent build is guaranteed to succeed and is
    /// the SINGLE evaluation of the argument. It mirrors <see cref="Probe"/>/<see cref="TryBuildOperand"/>
    /// exactly (same array-producing cases, same open-range refusal), so it is true iff the build succeeds as
    /// an array.
    /// </summary>
    public static bool IsArrayEligible(Expression expression) => Probe(expression).IsArray;

    // The pure-shape twin of TryBuildOperand: decides, WITHOUT evaluating any sub-expression, whether the
    // build would succeed (Succeeds — no refused open range on the eligible path) and whether the result is an
    // array (IsArray). Must track the builder's structure exactly so IsArrayEligible == (build result).
    private static (bool Succeeds, bool IsArray) Probe(Expression expression)
    {
        switch (expression)
        {
            case RangeReference:
                return (true, true);

            // An open/whole-column range in an array position is refused (the cost guard).
            case OpenRangeReference:
                return (false, false);

            case Row { Arguments: [RangeReference] }:
                return (true, true);

            case Row { Arguments: [OpenRangeReference] }:
                return (false, false);

            case BinaryOperation binary:
            {
                var left = Probe(binary.Left);
                if (!left.Succeeds)
                {
                    return (false, false);
                }

                var right = Probe(binary.Right);
                if (!right.Succeeds)
                {
                    return (false, false);
                }

                return (true, left.IsArray || right.IsArray);
            }

            case If ifNode when ifNode.Arguments.Length is 2 or 3:
            {
                var condition = Probe(ifNode.Arguments[0]);
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

                if (!Probe(ifNode.Arguments[1]).Succeeds)
                {
                    return (false, false);
                }

                if (ifNode.Arguments.Length == 3 && !Probe(ifNode.Arguments[2]).Succeeds)
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

    // ROW(range): every cell in row r shares the same worksheet row number (TopRow + r). Row-major.
    private sealed class RowNumbersOperand : ArrayOperand
    {
        private readonly int _topRow;
        private readonly int _rows;
        private readonly int _columns;

        public RowNumbersOperand(int topRow, int rows, int columns)
        {
            _topRow = topRow;
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

            return ComputedValue.Number(_topRow + index / _columns);
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

            // Whole-column / whole-row / one-sided open reference: the cost guard keeps it OUT of the
            // mini-CSE. Refuse so the whole evaluation reports "not an array".
            case OpenRangeReference:
                operand = null!;
                return false;

            case Row { Arguments: [RangeReference range] }:
                operand = new RowNumbersOperand(range.TopRow, range.RowCount, range.ColumnCount);
                return true;

            // ROW over an open range is likewise refused.
            case Row { Arguments: [OpenRangeReference] }:
                operand = null!;
                return false;

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

    private static ArrayOperand BuildRange(RangeReference range, EvaluationContext context)
    {
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(range.SheetName);

        return new RangeOperand(
            workbook,
            handle,
            range.SheetName,
            range.LeftColumn,
            range.TopRow,
            range.RowCount,
            range.ColumnCount
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
