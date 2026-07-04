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
/// <para>No production code consumes this yet — Phase B wires SUM/SMALL/LARGE/INDEX to it.</para>
/// </summary>
internal static class ArrayEvaluation
{
    /// <summary>
    /// Tries to evaluate <paramref name="expression"/> as an array. Returns <c>true</c> (with the vector in
    /// <paramref name="result"/>) only when the sub-tree genuinely produces an array — a closed range, a
    /// <c>ROW(range)</c>, or an operation/IF with at least one array operand. A bare scalar, a non-eligible
    /// node, or anything touching an open range returns <c>false</c> (the current scalar path is untouched).
    /// </summary>
    public static bool TryEvaluate(
        Expression expression,
        EvaluationContext context,
        out ArrayEvaluationResult result
    )
    {
        if (TryBuild(expression, context, out var operand) && operand.IsArray)
        {
            result = new ArrayEvaluationResult(operand.Rows, operand.Columns, operand.Values!);
            return true;
        }

        result = default;
        return false;
    }

    /// <summary>
    /// A CHEAP syntactic pre-check — no evaluation — for whether <paramref name="expression"/> would
    /// produce an array through the eligible structural set. Consumers gate on this BEFORE calling
    /// <see cref="TryEvaluate"/>, so the scalar hot path pays only a shallow type-walk and never a double
    /// evaluation: when this returns <c>false</c> the consumer keeps its existing scalar path untouched;
    /// when it returns <c>true</c> the subsequent <see cref="TryEvaluate"/> is guaranteed to succeed and
    /// is the SINGLE evaluation of the argument. It mirrors <see cref="Probe"/>/<see cref="TryBuild"/>
    /// exactly (same array-producing cases, same open-range refusal), so it is true iff TryEvaluate is.
    /// </summary>
    public static bool IsArrayEligible(Expression expression) => Probe(expression).IsArray;

    // The pure-shape twin of TryBuild: decides, WITHOUT evaluating any sub-expression, whether the build
    // would succeed (Succeeds — no refused open range on the eligible path) and whether the result is an
    // array (IsArray). Must track TryBuild's structure exactly so IsArrayEligible == (TryEvaluate result).
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

    /// <summary>
    /// A recursively-built operand: either a scalar (to broadcast) or a rectangular array. Building FAILS
    /// (returns <c>false</c> from <see cref="TryBuild"/>) only when the sub-tree reaches an open/whole-column
    /// range in an array position — the cost guard — so the whole evaluation degrades to "not an array".
    /// </summary>
    private readonly struct Operand
    {
        public bool IsArray { get; }
        public ComputedValue Scalar { get; }
        public int Rows { get; }
        public int Columns { get; }
        public ComputedValue[]? Values { get; }

        private Operand(bool isArray, ComputedValue scalar, int rows, int columns, ComputedValue[]? values)
        {
            IsArray = isArray;
            Scalar = scalar;
            Rows = rows;
            Columns = columns;
            Values = values;
        }

        public static Operand FromScalar(ComputedValue value) => new(false, value, 0, 0, null);

        public static Operand FromArray(int rows, int columns, ComputedValue[] values) =>
            new(true, default, rows, columns, values);

        /// <summary>
        /// The value at a row-major index within a target <paramref name="rows"/>×<paramref name="columns"/>
        /// shape: a scalar broadcasts to every position; an array of the SAME shape yields its element; an
        /// array of a DIFFERENT shape is a dimension mismatch and yields <c>#VALUE!</c> (Excel parity).
        /// </summary>
        public ComputedValue At(int index, int rows, int columns) =>
            !IsArray
                ? Scalar
                : Rows == rows && Columns == columns
                    ? Values![index]
                    : ComputedValue.Error(Error.Value);
    }

    private static bool TryBuild(Expression expression, EvaluationContext context, out Operand operand)
    {
        switch (expression)
        {
            case RangeReference range:
                operand = ExpandRange(range, context);
                return true;

            // Whole-column / whole-row / one-sided open reference: the cost guard keeps it OUT of the
            // mini-CSE. Refuse so the whole evaluation reports "not an array".
            case OpenRangeReference:
                operand = default;
                return false;

            case Row { Arguments: [RangeReference range] }:
                operand = ExpandRowNumbers(range);
                return true;

            // ROW over an open range is likewise refused.
            case Row { Arguments: [OpenRangeReference] }:
                operand = default;
                return false;

            case BinaryOperation binary:
                return TryBuildBinary(binary, context, out operand);

            case If ifNode when ifNode.Arguments.Length is 2 or 3:
                return TryBuildIf(ifNode, context, out operand);

            // Anything else is an opaque scalar: evaluate once and broadcast. (This is where nested scalar
            // functions — SUM(A:A), a bare cell, a literal — enter, without recursing into them.)
            default:
                operand = Operand.FromScalar(expression.Evaluate(context));
                return true;
        }
    }

    // Row-major expansion of a closed range's per-cell memoized values. Resolves the range origin and the sheet
    // handle ONCE (rather than re-parsing StartId/EndId and re-resolving the sheet per cell, as a per-cell
    // CellComputedValueAt would), then addresses each cell numerically through the dense accessor — identical
    // per-cell memoization, cycle guard and volatile taint. Iterating (row, column) keeps the documented
    // row-major layout for every shape (for the single-column/single-row K1 cases this equals
    // ExpandComputedValues' order).
    private static Operand ExpandRange(RangeReference range, EvaluationContext context)
    {
        var rows = range.RowCount;
        var columns = range.ColumnCount;
        var values = new ComputedValue[rows * columns];

        var originColumn = range.LeftColumn;
        var originRow = range.TopRow;
        var workbook = context.Workbook;
        var handle = workbook.ResolveDenseHandle(range.SheetName);

        var index = 0;
        for (var row = 1; row <= rows; row++)
        {
            for (var column = 1; column <= columns; column++)
            {
                values[index++] = workbook.GetCellValueDense(
                    handle,
                    range.SheetName,
                    originColumn + column - 1,
                    originRow + row - 1
                );
            }
        }

        return Operand.FromArray(rows, columns, values);
    }

    // ROW(range): every cell in row r shares the same worksheet row number (TopRow + r - 1). Row-major.
    private static Operand ExpandRowNumbers(RangeReference range)
    {
        var rows = range.RowCount;
        var columns = range.ColumnCount;
        var values = new ComputedValue[rows * columns];

        var index = 0;
        for (var row = 1; row <= rows; row++)
        {
            var rowNumber = ComputedValue.Number(range.TopRow + row - 1);
            for (var column = 1; column <= columns; column++)
            {
                values[index++] = rowNumber;
            }
        }

        return Operand.FromArray(rows, columns, values);
    }

    private static bool TryBuildBinary(BinaryOperation binary, EvaluationContext context, out Operand operand)
    {
        if (!TryBuild(binary.Left, context, out var left) || !TryBuild(binary.Right, context, out var right))
        {
            operand = default;
            return false;
        }

        // Neither side is an array → the whole operation is a scalar (broadcast), evaluated element-wise once.
        if (!left.IsArray && !right.IsArray)
        {
            operand = Operand.FromScalar(BinaryOperation.Apply(binary.Operator, left.Scalar, right.Scalar));
            return true;
        }

        var (rows, columns) = ResultShape(left, right);
        var values = new ComputedValue[rows * columns];

        for (var index = 0; index < values.Length; index++)
        {
            values[index] = BinaryOperation.Apply(
                binary.Operator,
                left.At(index, rows, columns),
                right.At(index, rows, columns)
            );
        }

        operand = Operand.FromArray(rows, columns, values);
        return true;
    }

    private static bool TryBuildIf(If ifNode, EvaluationContext context, out Operand operand)
    {
        if (!TryBuild(ifNode.Arguments[0], context, out var condition))
        {
            operand = default;
            return false;
        }

        // A scalar condition is not an array selection: treat the whole IF as an opaque scalar (its native
        // short-circuit Evaluate applies). Only an ARRAY condition drives the element-wise zip.
        if (!condition.IsArray)
        {
            operand = Operand.FromScalar(ifNode.Evaluate(context));
            return true;
        }

        if (!TryBuild(ifNode.Arguments[1], context, out var whenTrue))
        {
            operand = default;
            return false;
        }

        // IF without an else branch yields a logical FALSE where the condition is false (Excel), which the
        // aggregators then ignore — the whole point of the SMALL(IF(...)) idiom.
        var hasElse = ifNode.Arguments.Length == 3;
        var whenFalse = Operand.FromScalar(ComputedValue.Boolean(false));
        if (hasElse && !TryBuild(ifNode.Arguments[2], context, out whenFalse))
        {
            operand = default;
            return false;
        }

        var rows = condition.Rows;
        var columns = condition.Columns;
        var values = new ComputedValue[rows * columns];

        for (var index = 0; index < values.Length; index++)
        {
            var conditionValue = condition.Values![index];

            // A condition that is (or coerces from) an error propagates that error at this position.
            if (conditionValue.CoerceToBool(out var taken) is { } error)
            {
                values[index] = ComputedValue.Error(error);
                continue;
            }

            values[index] = taken
                ? whenTrue.At(index, rows, columns)
                : whenFalse.At(index, rows, columns);
        }

        operand = Operand.FromArray(rows, columns, values);
        return true;
    }

    // The shape of a binary result: a scalar takes the other side's shape; two equal-shaped arrays keep it;
    // mismatched arrays produce the per-axis maximum, filled entirely with #VALUE! by the At() mismatch rule.
    private static (int Rows, int Columns) ResultShape(in Operand left, in Operand right)
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
}
