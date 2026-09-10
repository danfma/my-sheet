namespace Danfma.MySheet.Expressions;

/// <summary>
/// The shape pieces shared by every <see cref="IArrayProducer"/> (Phase 7: <c>FILTER</c>, <c>SORT</c>,
/// <c>UNIQUE</c>, <c>SEQUENCE</c>) — and the HARD INVARIANT their operands live under.
/// </summary>
/// <remarks>
/// <para><b>Invariant.</b> An <see cref="ArrayOperand"/> handed back by
/// <see cref="IArrayProducer.TryBuildArrayOperand"/> has <c>IsArray</c> and
/// <c>Rows &gt;= 1 &amp;&amp; Columns &gt;= 1</c>. An EMPTY result — nothing kept by <c>FILTER</c>, nothing
/// left by <c>UNIQUE(…, exactly_once)</c> — is a 1x1 <see cref="SingletonArrayOperand"/> carrying the error
/// (<c>#CALC!</c>, once the error set has it; a bad <c>SEQUENCE</c> size is a 1x1 <c>#VALUE!</c> the same
/// way), NEVER a 0-extent array. A 1x1 SOURCE is likewise wrapped in a singleton, because
/// <see cref="ScalarOperand"/> reports 0x0 and can never be a result (blocker B1).</para>
///
/// <para><b>What breaks if it is violated</b> (measured on this tree, pinned in
/// <c>ArrayProducerContractTests</c>): <c>ArrayEvaluation.ArrayStream.Length</c> is <c>Rows * Columns</c>
/// and has NO emptiness channel, so a 0x1 operand enumerates nothing and every consumer answers as if the
/// argument were an empty range — <c>SUM</c> 0, <c>COUNT</c> 0, <c>AVERAGE</c> <c>#DIV/0!</c> — where the
/// oracle answers <c>#CALC!</c> for the empty <c>FILTER</c> (<c>SUM(FILTER(A1:A3,A1:A3&gt;100))</c>,
/// Aspose.Cells 26.6.0, 2026-09-10, both entry modes). A silent 0 in place of an error is the exact class
/// of wrong answer this phase must not ship. The 1x1 shape is also what makes an empty result COMPOSE:
/// under an operator it broadcasts (Phase 10 rule 1) and fills every element with its error —
/// <c>SUM(FILTER(A1:A3,A1:A3&gt;100,7)*A1:A3)</c> is 98, the singleton's 7 against every cell — whereas a
/// 0-row extent would have to be folded as a SHAPE by <c>ArrayEvaluation.ShapeFold</c>, which now assumes
/// it never sees one.</para>
///
/// <para><b>Enforced, not aspirational.</b> <see cref="RequireProducerShape"/> throws at construction —
/// once per build, never per element, in every configuration (a <c>Debug.Assert</c> is compiled out of the
/// Release build the gates run) — and <see cref="AxisSelectionOperand"/> calls it; every producer-owned
/// operand class (<c>SequenceOperand</c>) must call it too. <see cref="SingletonArrayOperand"/> is 1x1 by
/// construction. <c>ArrayProducerContractTests</c> violates the rule and fails.</para>
/// </remarks>
internal static class ArrayShaping
{
    /// <summary>
    /// Throws unless <paramref name="rows"/> and <paramref name="columns"/> are both at least 1 — the
    /// invariant every producer operand is constructed under (see the type remarks). The exception names
    /// the rule, so a violating producer fails loudly at build time instead of streaming nothing.
    /// </summary>
    public static void RequireProducerShape(int rows, int columns)
    {
        if (rows < 1 || columns < 1)
        {
            throw new InvalidOperationException(
                $"A producer's array operand must satisfy Rows >= 1 && Columns >= 1, got {rows}x{columns}: "
                    + "an empty result is a 1x1 SingletonArrayOperand carrying the error, never a 0-extent array."
            );
        }
    }
}

/// <summary>
/// Which axis of a source rectangle an <see cref="AxisSelectionOperand"/> selects along: <c>FILTER</c>
/// keeps rows (a column include) or columns (a row include); <c>SORT</c>/<c>UNIQUE</c> follow
/// <c>by_col</c>.
/// </summary>
internal enum ArrayAxis
{
    Rows,
    Columns,
}

/// <summary>
/// A 1x1 ARRAY holding one value — the shape of every producer result that is not a rectangle: a 1x1
/// source (<c>SUM(SORT(A1))</c> = 5), an <c>if_empty</c> value, and every error a producer answers for
/// itself (a bad argument, an empty result). It exists because <see cref="ScalarOperand"/> has
/// <c>IsArray => false</c> and 0x0 and therefore can never be a RESULT, only a broadcast operand.
/// </summary>
/// <remarks>
/// A 1x1 operand covers EVERY position of any extent (<see cref="Broadcasting"/> rule 1 — both axes are
/// of extent 1, so <see cref="Broadcasting.TryProject"/> always answers index 0): the projection is the
/// identity onto the one value, which is why <see cref="At"/> returns it directly. That is Excel's rule
/// and the reason <c>FILTER(A1,TRUE)*A1:A3</c> multiplies every cell (70) and an empty <c>FILTER</c> under
/// an operator errors every element (Aspose.Cells 26.6.0, 2026-09-10, CSE column;
/// <c>ArrayProducerContractTests</c>) — do not "fix" the broadcast.
/// </remarks>
internal sealed class SingletonArrayOperand : ArrayOperand
{
    private readonly ComputedValue _value;

    public SingletonArrayOperand(ComputedValue value) => _value = value;

    public override bool IsArray => true;
    public override int Rows => 1;
    public override int Columns => 1;

    public override ComputedValue At(int index, int rows, int columns) => _value;
}

/// <summary>
/// A selection or permutation of ONE axis of a source array — the single operand behind <c>FILTER</c>,
/// <c>SORT</c> and <c>UNIQUE</c>, which differ only in how they compute <c>selection</c> (kept positions,
/// a sort permutation, first appearances) at BUILD time. The values stay on demand: each read maps this
/// operand's own (row, column) through <c>selection</c> on the chosen axis and pulls that element from the
/// source at the SOURCE's extent, so a composite source (<c>A1:A3*2</c>, another producer) is read exactly
/// as it would be anywhere else in the tree.
/// </summary>
/// <remarks>
/// Its shape is the selection's length on the chosen axis and the source's extent on the other, and it
/// projects through <see cref="Broadcasting.TryProject"/> before anything else, like every array operand:
/// a position the selection does not cover is <c>#N/A</c> — <c>SUM(FILTER(A1:A3,A1:A3&gt;0)*B1:B3)</c> is
/// <c>#N/A</c> with <c>COUNT</c> 2 (a 2x1 selection against a 3x1 range; Aspose.Cells 26.6.0, 2026-09-10,
/// CSE column). The constructor enforces <see cref="ArrayShaping"/>'s invariant (a non-empty selection over
/// an ARRAY source) and rejects an index past the source's axis, so a producer bug fails at build rather
/// than as a stray exception per element.
/// </remarks>
internal sealed class AxisSelectionOperand : ArrayOperand
{
    private readonly ArrayOperand _source;
    private readonly ArrayAxis _axis;
    private readonly int[] _selection;
    private readonly int _rows;
    private readonly int _columns;

    public AxisSelectionOperand(ArrayOperand source, ArrayAxis axis, int[] selection)
    {
        if (!source.IsArray)
        {
            throw new InvalidOperationException(
                "An AxisSelectionOperand's source must be an array: a scalar source reports 0x0 (ScalarOperand) "
                    + "and must be wrapped in a SingletonArrayOperand by the producer first (blocker B1)."
            );
        }

        _source = source;
        _axis = axis;
        _selection = selection;

        (_rows, _columns) =
            axis is ArrayAxis.Rows
                ? (selection.Length, source.Columns)
                : (source.Rows, selection.Length);

        ArrayShaping.RequireProducerShape(_rows, _columns);

        var extent = axis is ArrayAxis.Rows ? source.Rows : source.Columns;

        foreach (var index in selection)
        {
            if ((uint)index >= (uint)extent)
            {
                throw new InvalidOperationException(
                    $"An AxisSelectionOperand selection index {index} is outside the source's {axis} extent of {extent}."
                );
            }
        }
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

        var (sourceRow, sourceColumn) =
            _axis is ArrayAxis.Rows ? (_selection[row], column) : (row, _selection[column]);

        return _source.At(
            sourceRow * _source.Columns + sourceColumn,
            _source.Rows,
            _source.Columns
        );
    }
}
