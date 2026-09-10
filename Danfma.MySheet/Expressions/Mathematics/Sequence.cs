using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

/// <summary>
/// <c>SEQUENCE(rows, [columns], [start], [step])</c> — the first mini-CSE PRODUCER (Phase 7), and the only one
/// with no input array: a <paramref name="rows"/>x<c>columns</c> array filled row-major with
/// <c>start + i * step</c>. It reaches every consumer through <see cref="IArrayProducer"/> —
/// <c>SUM(SEQUENCE(5))</c> is 15, <c>INDEX(SEQUENCE(2,3),2,2)</c> is 5 — and in a cell it answers its
/// top-left element (<see cref="ArrayEvaluation.FirstElement"/>): <c>=SEQUENCE(2,3,7,1)</c> is 7.
/// </summary>
/// <remarks>
/// <para><b>Arguments</b>, each evaluated ONCE at build time (the read the laziness contract sanctions). An
/// OMITTED optional — absent, or the parser's <see cref="BlankValue"/> for an empty slot — defaults to 1
/// (<c>SUM(SEQUENCE(2,2,,))</c> = 10, <c>=SEQUENCE(2,,,)</c> = 1); a blank CELL is a value and coerces to 0
/// like everywhere else (<c>SUM(SEQUENCE(2,2,A9))</c> = 6 with A9 empty, <c>SUM(SEQUENCE(A9))</c> =
/// <c>#VALUE!</c>). <c>rows</c> and <c>columns</c> truncate (<c>SUM(SEQUENCE(2.7))</c> = 3,
/// <c>SUM(SEQUENCE(2,2.9))</c> = 10); <c>start</c> and <c>step</c> are kept as given
/// (<c>SUM(SEQUENCE(1,1,1.5,0.25))</c> = 1.5), a zero or negative step included.</para>
///
/// <para><b>Errors</b> are the producer's own 1x1 <see cref="SingletonArrayOperand"/>, never a refused
/// build, so <c>ProbeArray</c> and <c>TryBuildArrayOperand</c> stay in lockstep (<c>(true, true)</c> /
/// <c>true</c>, always). A coercion error passes through unchanged (<c>SUM(SEQUENCE(1/0))</c> =
/// <c>#DIV/0!</c>, <c>SUM(SEQUENCE("x"))</c> = <c>#VALUE!</c>); a size below 1 after truncation is
/// <c>#VALUE!</c> — <c>SUM(SEQUENCE(0))</c>, <c>SUM(SEQUENCE(-1))</c>, <c>=SEQUENCE(0,0)</c>,
/// <c>SUM(SEQUENCE(0.5))</c>, with <c>ERROR.TYPE(SEQUENCE(-1))</c> = 3 — NOT the phase design's
/// <c>#CALC!</c> (correction B2). Every number in this remark was measured on Aspose.Cells 26.6.0,
/// 2026-09-10, plain and CSE entry agreeing; pinned in <c>SequenceTests</c> and <c>DynamicArrayTests</c>.</para>
///
/// <para><b>The size cap is MySheet's own deviation.</b> <c>rows &gt; 1,048,576</c>, <c>columns &gt; 16,384</c>
/// or <c>rows * columns &gt; 1,048,576</c> answers a 1x1 <c>#NUM!</c>. The oracle has NO such cap in a consumed
/// position — <c>ROWS(SEQUENCE(1048577))</c> = 1048577, <c>COLUMNS(SEQUENCE(1,16385))</c> = 16385,
/// <c>SUM(SEQUENCE(1048577))</c> = 549757386753 (Aspose.Cells 26.6.0, 2026-09-10, both modes) — the cap
/// exists because <c>ArrayStream</c> is lazy but every consumer iterates every element, so
/// <c>SEQUENCE(1e6,1e4)</c> would otherwise hang the consumer. It is pinned as a labelled divergence
/// (<c>DynamicArrayTests.Sequence_BeyondTheGrid_IsNumError_TheOnePinnedDivergence</c>).</para>
/// </remarks>
[MemoryPackable]
public sealed partial record Sequence(Expression[] Arguments) : Function, IArrayProducer
{
    // The grid's extents, and the cell budget: one full column's worth of elements.
    private const double MaxRows = 1_048_576;
    private const double MaxColumns = 16_384;
    private const double MaxCells = 1_048_576;

    public override ComputedValue Evaluate(EvaluationContext context) =>
        ArrayEvaluation.FirstElement(this, context);

    // No child can be refused (there is no array argument), and every failure is a 1x1 error operand, so
    // the build below never returns false — which is exactly what this probe promises.
    (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
        (true, true);

    bool IArrayProducer.TryBuildArrayOperand(EvaluationContext context, out ArrayOperand operand)
    {
        operand = Build(context);
        return true;
    }

    private ArrayOperand Build(EvaluationContext context)
    {
        if (ReadArgument(0, context, out var rows) is { } rowsError)
        {
            return Singleton(rowsError);
        }

        if (ReadArgument(1, context, out var columns) is { } columnsError)
        {
            return Singleton(columnsError);
        }

        if (ReadArgument(2, context, out var start) is { } startError)
        {
            return Singleton(startError);
        }

        if (ReadArgument(3, context, out var step) is { } stepError)
        {
            return Singleton(stepError);
        }

        rows = Math.Truncate(rows);
        columns = Math.Truncate(columns);

        // Written as the negation of the accepted range so that a NaN size lands here too, instead of
        // reaching the int casts below.
        if (!(rows >= 1 && columns >= 1))
        {
            return Singleton(Error.Value);
        }

        if (rows > MaxRows || columns > MaxColumns || rows * columns > MaxCells)
        {
            return Singleton(Error.Num);
        }

        return new SequenceOperand((int)rows, (int)columns, start, step);
    }

    // An omitted optional (absent, or the parser's BlankValue for an empty slot) is 1; anything else is a
    // value and coerces as a number — a blank cell to 0, TRUE to 1, "3" to 3, "x" to #VALUE!.
    private Error? ReadArgument(int index, EvaluationContext context, out double value)
    {
        if (index >= Arguments.Length || Arguments[index] is BlankValue)
        {
            value = 1;
            return null;
        }

        return Arguments[index].Evaluate(context).CoerceToNumber(out value);
    }

    private static SingletonArrayOperand Singleton(Error error) => new(ComputedValue.Error(error));
}

/// <summary>
/// The operand behind a valid <see cref="Sequence"/>: <c>start + i * step</c> at row-major position
/// <c>i</c>, computed on demand, never materialized. Internal (not <c>file</c>-scoped) so that the
/// invariant guard in its constructor can be exercised directly by <c>SequenceTests</c>.
/// </summary>
/// <remarks>
/// It projects through <see cref="Broadcasting.TryProject"/> before reading, like every array operand:
/// an uncovered position is <c>#N/A</c>, an axis of extent 1 repeats (<c>SUM(SEQUENCE(3)*B1:B3)</c> = 14,
/// Aspose.Cells 26.6.0, 2026-09-10, CSE column; plain entry is <c>#VALUE!</c>). The constructor enforces
/// <see cref="ArrayShaping"/>'s invariant — <c>Rows &gt;= 1 &amp;&amp; Columns &gt;= 1</c> — because
/// nothing else does for an operand class outside <c>ArrayShaping.cs</c>: a 0-extent operand would
/// stream nothing and make <c>SUM</c> answer 0 in silence.
/// </remarks>
internal sealed class SequenceOperand : ArrayOperand
{
    private readonly int _rows;
    private readonly int _columns;
    private readonly double _start;
    private readonly double _step;

    public SequenceOperand(int rows, int columns, double start, double step)
    {
        ArrayShaping.RequireProducerShape(rows, columns);

        _rows = rows;
        _columns = columns;
        _start = start;
        _step = step;
    }

    public override bool IsArray => true;
    public override int Rows => _rows;
    public override int Columns => _columns;

    public override ComputedValue At(int index, int rows, int columns) =>
        Broadcasting.TryProject(index, rows, columns, _rows, _columns, out var own)
            ? ComputedValue.Number(_start + own * _step)
            : ComputedValue.Error(Error.NA);
}
