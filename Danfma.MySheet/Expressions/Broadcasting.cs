using System.Runtime.CompilerServices;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Excel's rule for array operands of DIFFERENT shapes inside the mini-CSE (<see cref="ArrayEvaluation"/>),
/// in two halves: <see cref="Axis"/> decides the extent of a combined result at BUILD time (folded by
/// <c>ArrayEvaluation.ShapeFold</c>), and <see cref="TryProject"/> maps a position of that extent to an
/// operand's own element at READ time (every array operand's <see cref="ArrayOperand.At"/>).
/// </summary>
/// <remarks>
/// The rule, per axis independently, as measured on Aspose.Cells 26.6.0 (2026-09-10, CSE-entered — the
/// column the mini-CSE reproduces; pinned in <c>VectorBroadcastingTests</c>):
/// <list type="number">
/// <item>a 1x1 range broadcasts everywhere, like a scalar — <c>SUM(A1:C3*E1:E1)</c> = 45;</item>
/// <item>a 1xN row repeats down every row — <c>SUM(A1:C3*E5:G5)</c> = 960;</item>
/// <item>an Nx1 column repeats across every column — <c>SUM(A1:C3*E1:E3)</c> = 108;</item>
/// <item>an Nx1 against a 1xM is the NxM outer product — <c>SUM(E1:E3*E5:G5)</c> = 360;</item>
/// <item>where BOTH extents exceed 1 and differ, the result takes the LARGER, and every position the
/// shorter operand does not cover is a per-element <c>#N/A</c> — <c>SUM(A1:C3*H1:H2)</c> is <c>#N/A</c>
/// while <c>COUNT(A1:C3*H1:H2)</c> is 6 and <c>IFERROR</c> recovers the three uncovered elements.</item>
/// </list>
/// The <c>#N/A</c> is answered by the OPERAND that is uncovered, whichever kind it is: a leaf
/// (<c>RangeOperand</c>, <c>PositionNumbersOperand</c>), or a COMPOSITE uncovered at its consumer's extent,
/// which answers its own <c>#N/A</c> and never asks its children — the four composites in
/// <c>ArrayOperands.cs</c> (<c>BinaryOperand</c>, <c>IfOperand</c>, <c>UnaryOperand</c>,
/// <c>LiftedFunctionOperand</c>) each project before descending, so <c>(A1:B2*1)</c> read at 3x3 is the
/// composite's own <c>#N/A</c> (<c>VectorBroadcastingTests</c>). The operator's existing left-then-right
/// error precedence then decides between one side's own error and the other side's uncovered position
/// (<c>INDEX(A1:C3*H1:H2,3,1)</c> with <c>A3</c> = <c>#DIV/0!</c> is <c>#DIV/0!</c>, the operands swapped
/// it is <c>#N/A</c>).
/// </remarks>
internal static class Broadcasting
{
    /// <summary>
    /// The combined extent of two operands along one axis: an extent of 1 defers to the other operand;
    /// otherwise the larger extent wins, and the shorter operand marks what it does not cover at read time.
    /// </summary>
    public static int Axis(int a, int b) =>
        a == 1 ? b
        : b == 1 ? a
        : Math.Max(a, b);

    /// <summary>
    /// Projects the consumer's row-major <paramref name="index"/> within a <paramref name="rows"/>×
    /// <paramref name="columns"/> extent onto an operand of <paramref name="ownRows"/>×
    /// <paramref name="ownColumns"/>: <c>true</c> with the operand's own row-major index, or <c>false</c>
    /// when the operand does not cover that position. Equal shapes are two comparisons and the identity —
    /// no division on the path every equal-shaped formula takes; otherwise the index is decomposed once, an
    /// axis of extent 1 forces its coordinate to 0, and a coordinate at or beyond the operand's extent is
    /// uncovered. Integer arithmetic on the stack: no allocation per element.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryProject(
        int index,
        int rows,
        int columns,
        int ownRows,
        int ownColumns,
        out int ownIndex
    )
    {
        if (ownRows == rows && ownColumns == columns)
        {
            ownIndex = index;
            return true;
        }

        var row = index / columns;
        var column = index % columns;

        if (ownRows == 1)
        {
            row = 0;
        }
        else if (row >= ownRows)
        {
            ownIndex = -1;
            return false;
        }

        if (ownColumns == 1)
        {
            column = 0;
        }
        else if (column >= ownColumns)
        {
            ownIndex = -1;
            return false;
        }

        ownIndex = row * ownColumns + column;
        return true;
    }
}
