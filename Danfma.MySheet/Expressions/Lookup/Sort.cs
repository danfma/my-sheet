using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// <c>SORT(array, [sort_index], [sort_order], [by_col])</c> — the rows (or, with <c>by_col</c>, the columns)
/// of <c>array</c> permuted by one key column (or row), as a mini-CSE producer (<see cref="IArrayProducer"/>):
/// <c>INDEX(SORT(A1:A3),1)</c> is 0 over 5, 0, 9, <c>INDEX(SORT(A1:A3,1,-1),1)</c> is 9, and a cell shows the
/// top-left (<see cref="ArrayEvaluation.FirstElement"/>: <c>=SORT(A1:B3,1,-1)</c> is 9). The KEYS are read
/// once at build time and the permutation fixed then; the values stay on demand through
/// <see cref="AxisSelectionOperand"/>, so a blank arrives as a blank.
/// </summary>
/// <remarks>
/// <para><b>Arguments</b>, evaluated once at build, in order (the first error wins: <c>SORT(A1:A3,1,1/0)</c>
/// and <c>SORT(A1:A3,1,1,1/0)</c> are <c>#DIV/0!</c>). <c>sort_index</c> defaults to 1, truncates
/// (<c>1.9</c> → 1, <c>"1"</c> and TRUE → 1) and must fall within the extent ACROSS the sorted axis —
/// <c>SORT(A1:A3,0)</c>, <c>SORT(A1:A3,2)</c>, <c>SORT(A1:B3,3)</c>, <c>SORT(A1:B3,4,1,TRUE)</c>,
/// <c>SORT(A1:A3,"x")</c> and <c>SORT(A1:A3,A9)</c> (a blank CELL is 0) are <c>#VALUE!</c>, while
/// <c>SORT(A1:B3,3,1,TRUE)</c> sorts by the third ROW. <c>sort_order</c> defaults to 1, truncates
/// (<c>-1.5</c> → -1, <c>"-1"</c> → -1, TRUE → 1) and must then be 1 or -1: <c>0</c>, <c>2</c> and
/// <c>-0.5</c> are <c>#VALUE!</c>. <c>by_col</c> defaults to FALSE and coerces like every flag
/// (<see cref="SelectionProducers"/>: <c>"TRUE"</c> and 2 are TRUE, a blank cell is FALSE, <c>"x"</c> is
/// <c>#VALUE!</c>). An EMPTY <c>sort_index</c> slot is omitted (<c>INDEX(SORT(A1:A3,,-1),1)</c> = 9); an empty
/// <c>by_col</c> slot is FALSE either way. An empty, blank-cell or text <c>sort_order</c> is NOT measured:
/// Aspose.Cells 26.6.0 throws a <c>NullReferenceException</c> from <c>CalculateFormula</c> for
/// <c>SORT(A1:A3,1,)</c>, <c>SORT(A1:A3,1,A9)</c> and <c>SORT(A1:A3,1,"x")</c>, so those three follow the
/// rule above — omitted → 1, blank cell → 0 → <c>#VALUE!</c>, text → coercion <c>#VALUE!</c> — on Microsoft's
/// page alone.</para>
///
/// <para><b>Order.</b> Values compare through <see cref="ValueCoercion.Compare"/> — number &lt; text &lt;
/// FALSE &lt; TRUE, text case-insensitively — negated for descending; ties keep the SOURCE order in BOTH
/// directions (<c>SORT(N1:O4)</c> tags a, c, b, d and <c>SORT(N1:O4,1,-1)</c> b, d, a, c — a stable sort each
/// way, not a reversal; <c>SORT(C1:C3)</c> over "a", "A", "b" is a, A, b and descending b, a, A). Two kinds
/// are NOT what <c>Compare</c> would do on its own, and the oracle contradicts the phase design on both:
/// BLANKS go last in both directions (<c>SORT(A5:A8)</c> over 7, blank, "t", 7 is 7, 7, "t", blank and
/// descending "t", 7, 7, blank; <c>ISBLANK</c> of the last row TRUE) instead of sorting as 0, and ERRORS are
/// sorted rather than propagated — after every value ascending, before every value descending, in their
/// source order among themselves (<c>SORT(W1:W5)</c> over 2, <c>#DIV/0!</c>, 3, <c>#N/A</c>, <c>#DIV/0!</c>
/// is 2, 3, <c>#DIV/0!</c>, <c>#N/A</c>, <c>#DIV/0!</c> and descending <c>#DIV/0!</c>, <c>#N/A</c>,
/// <c>#DIV/0!</c>, 3, 2). <c>SUM(SORT(E1:E3))</c> is <c>#DIV/0!</c> only because <c>SUM</c> propagates what
/// it is handed.</para>
///
/// <para><b>Refusal.</b> The build returns <c>false</c> exactly when the array's own build is refused (an
/// open range), and <c>ProbeArray</c> asks the same question, so the two cannot drift; every other failure
/// is the producer's own 1x1 <see cref="SingletonArrayOperand"/>. A 1x1 source is a 1x1 result
/// (<c>SUM(SORT(A1))</c> = 5, <c>=SORT("x")</c> = "x", <c>SUM(SORT(1/0))</c> = <c>#DIV/0!</c>).</para>
///
/// <para>Every number above was measured on Aspose.Cells 26.6.0, 2026-09-10, plain and array-entered entry
/// agreeing unless named; pinned in <c>SortTests</c> and <c>DynamicArrayTests</c>. One split is deliberately
/// NOT pinned: <c>INDEX(SORT(A1:A3,1/0),1)</c> is <c>#DIV/0!</c> plain and <c>#VALUE!</c> array-entered on
/// the oracle; this implementation propagates the argument's own error like the other two slots do.</para>
/// </remarks>
[MemoryPackable]
public sealed partial record Sort(Expression[] Arguments) : Function, IArrayProducer
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ArrayEvaluation.FirstElement(this, context);

    (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
        SelectionProducers.ProbeSource(Arguments[0], context) ? (true, true) : (false, false);

    bool IArrayProducer.TryBuildArrayOperand(EvaluationContext context, out ArrayOperand operand)
    {
        if (!SelectionProducers.TryBuildSource(Arguments[0], context, out var source))
        {
            operand = null!;
            return false;
        }

        operand = Build(source, context);
        return true;
    }

    private ArrayOperand Build(ArrayOperand source, EvaluationContext context)
    {
        if (ReadNumber(1, context, out var sortIndex) is { } indexError)
        {
            return SelectionProducers.Singleton(indexError);
        }

        if (ReadNumber(2, context, out var sortOrder) is { } orderError)
        {
            return SelectionProducers.Singleton(orderError);
        }

        if (
            SelectionProducers.ReadFlag(Arguments, 3, context, out var byColumn) is
            { } byColumnError
        )
        {
            return SelectionProducers.Singleton(byColumnError);
        }

        var axis = byColumn ? ArrayAxis.Columns : ArrayAxis.Rows;

        sortIndex = Math.Truncate(sortIndex);
        sortOrder = Math.Truncate(sortOrder);

        // Written as the negation of the accepted range so that a NaN lands here too.
        if (!(sortIndex >= 1 && sortIndex <= SelectionProducers.CrossExtent(source, axis)))
        {
            return SelectionProducers.Singleton(Error.Value);
        }

        if (sortOrder is not (1 or -1))
        {
            return SelectionProducers.Singleton(Error.Value);
        }

        var length = SelectionProducers.Extent(source, axis);
        var keyOffset = (int)sortIndex - 1;
        var keys = new ComputedValue[length];

        for (var i = 0; i < length; i++)
        {
            keys[i] = SelectionProducers.ElementAt(source, axis, i, keyOffset);
        }

        var permutation = SelectionProducers.Everything(source, axis);

        Array.Sort(permutation, new KeyOrder(keys, descending: sortOrder < 0));

        return new AxisSelectionOperand(source, axis, permutation);
    }

    // An omitted optional (absent, or the parser's BlankValue for an empty slot) is 1; anything else is a
    // value and coerces as a number — a blank cell to 0, TRUE to 1, "-1" to -1, "x" to #VALUE!.
    private Error? ReadNumber(int index, EvaluationContext context, out double value)
    {
        if (SelectionProducers.IsOmitted(Arguments, index))
        {
            value = 1;
            return null;
        }

        return Arguments[index].Evaluate(context).CoerceToNumber(out value);
    }

    // The order over positions: by key (see the type remarks), then by position — the explicit tie-break
    // that makes Array.Sort's introsort stable in both directions.
    private sealed class KeyOrder(ComputedValue[] keys, bool descending) : IComparer<int>
    {
        public int Compare(int left, int right)
        {
            var order = CompareKeys(keys[left], keys[right]);

            return order != 0 ? order : left.CompareTo(right);
        }

        private int CompareKeys(in ComputedValue left, in ComputedValue right)
        {
            // Blanks last, whichever the direction — so they are placed before the direction is applied.
            var leftBlank = left.Kind == ComputedValueKind.Blank;
            var rightBlank = right.Kind == ComputedValueKind.Blank;

            if (leftBlank || rightBlank)
            {
                return leftBlank == rightBlank ? 0
                    : leftBlank ? 1
                    : -1;
            }

            // Errors rank above every value and equal to each other, then the direction applies: last
            // ascending, first descending, source order among themselves either way.
            var leftError = left.Kind == ComputedValueKind.Error;
            var rightError = right.Kind == ComputedValueKind.Error;

            var order =
                leftError || rightError
                    ? leftError == rightError
                        ? 0
                        : leftError
                            ? 1
                            : -1
                    : ValueCoercion.Compare(left, right);

            return descending ? -order : order;
        }
    }
}
