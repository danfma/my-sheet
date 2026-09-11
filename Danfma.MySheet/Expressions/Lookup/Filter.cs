using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// <c>FILTER(array, include, [if_empty])</c> — the rows (or columns) of <c>array</c> whose <c>include</c>
/// element is TRUE, as a mini-CSE producer (<see cref="IArrayProducer"/>): <c>SUM(FILTER(A1:A3,A1:A3&gt;0))</c>
/// is 14 over 5, 0, 9 and <c>INDEX(FILTER(A1:A3,A1:A3&gt;0),2)</c> is 9; in a cell it answers its top-left
/// element (<see cref="ArrayEvaluation.FirstElement(Expression, EvaluationContext)"/>). The include is read ONCE at build time — the
/// data-dependent length is fixed then — and the kept VALUES stay on demand through
/// <see cref="AxisSelectionOperand"/>, so a blank survives the selection as a blank
/// (<c>COUNTA(FILTER(A5:A8,A5:A8&lt;&gt;"zzz"))</c> = 3 = <c>COUNTA(A5:A8)</c>, <c>ISBLANK</c> of the second
/// row TRUE) with no normalization anywhere.
/// </summary>
/// <remarks>
/// <para><b>Axis.</b> The include must match the array's HEIGHT as a single column (rows are kept) or its
/// WIDTH as a single row (columns are kept); anything else is a 1x1 <c>#VALUE!</c> —
/// <c>SUM(FILTER(A1:A3,B1:B2&gt;0))</c>, <c>SUM(FILTER(A1:A3,A1:B3&gt;0))</c>, <c>SUM(FILTER(A1:B3,A1:C1&gt;0))</c>.
/// A SCALAR include is a 1x1 include under the same rule, and so is a 1x1 array: it matches a column or a row
/// (<c>SUM(FILTER(A1:A3,TRUE))</c> = 14, <c>SUM(FILTER(A1:C1,TRUE))</c> = 6, <c>SUM(FILTER(A1:A3,B1:B1&gt;0))</c>
/// = 14) but NOT a rectangle (<c>SUM(FILTER(A1:B3,TRUE))</c> and <c>SUM(FILTER(A1:B3,SEQUENCE(1)))</c> are
/// <c>#VALUE!</c>) — so correction B1's "a scalar include broadcasts" holds for a vector source only.</para>
///
/// <para><b>Coercion.</b> Each include element is coerced with the text words accepted
/// (<see cref="ValueCoercion.CoerceToBoolAllowingTextWords"/>, the rule Phase 11b measured for this slot):
/// over cells <c>"TRUE"</c>, <c>"FALSE"</c>, <c>"true"</c> the filter keeps rows 1 and 3. The two clauses of
/// Microsoft's "an error or cannot be converted to a Boolean" split on the oracle: an ERROR element is the
/// producer's whole answer (<c>SUM(FILTER(A1:A3,E1:E3&gt;0))</c> = <c>#DIV/0!</c>), while an element that
/// cannot be converted is simply NOT KEPT — <c>SUM(FILTER(A1:A3,F1:F3))</c> over TRUE, <c>"x"</c>, TRUE is 14
/// with <c>COUNT</c> 2, and over <c>" TRUE "</c>, <c>"yes"</c>, <c>"1"</c> nothing is kept and the answer is
/// the empty result's <c>#CALC!</c>, not a coercion error. A ONE-element include is the exception: there the
/// unconvertible text IS <c>#VALUE!</c>, even with <c>if_empty</c> (<c>SUM(FILTER(A1:A3,"x"))</c>,
/// <c>FILTER(A1:A3," TRUE ","none")</c>, <c>FILTER(A1:A3,F2:F2,"none")</c>), and a falsy one keeps nothing
/// as <c>#VALUE!</c> rather than <c>#CALC!</c> (<c>SUM(FILTER(A1:A3,FALSE))</c>, <c>(…,0)</c>, <c>(…,A9)</c>,
/// <c>(…,B1:B1&gt;5)</c>) unless <c>if_empty</c> answers (<c>FILTER(A1:A3,FALSE,"none")</c> = "none").</para>
///
/// <para><b>Empty result.</b> Nothing kept is a 1x1 <c>#CALC!</c> (<c>SUM(FILTER(A1:A3,A1:A3&gt;100))</c>,
/// <c>ERROR.TYPE</c> 14, <c>COUNT</c> 0) — never a 0-extent array (<see cref="ArrayShaping"/>). With a third
/// argument the result is that argument instead, built only then: a scalar as a 1x1 array
/// (<c>FILTER(A1:A3,A1:A3&gt;100,"none")</c> = "none", <c>COUNTA</c> 1), an ARRAY as that array
/// (<c>SUM(FILTER(A1:A3,A1:A3&gt;100,B1:B3))</c> = 6, <c>ROWS</c> 3), an error as itself
/// (<c>FILTER(A1:A3,A1:A3&gt;100,1/0)</c> = <c>#DIV/0!</c>, while <c>SUM(FILTER(A1:A3,A1:A3&gt;0,1/0))</c> = 14
/// because it is never read). An EMPTY <c>if_empty</c> slot is a blank VALUE, not an omitted argument —
/// <c>FILTER(A1:A3,A1:A3&gt;100,)</c> is a 1x1 blank (<c>ISBLANK</c> TRUE, <c>COUNTA</c> 0, <c>ROWS</c> 1) —
/// which is why this slot alone does not follow <see cref="SelectionProducers"/>' omitted rule.</para>
///
/// <para><b>Refusal.</b> The build returns <c>false</c> exactly when the array's or the include's own build
/// is refused (an open range: <c>SUM(FILTER(A:A,A:A&gt;0))</c> is <c>#VALUE!</c>, the phase's pinned
/// deviation from the oracle's 28 array-entered), and <c>ProbeArray</c> probes those same two arguments, so
/// probe and build cannot drift. <c>if_empty</c> is NOT probed, because it is built lazily; a refused
/// <c>if_empty</c> is therefore a 1x1 <c>#VALUE!</c> from the build rather than a refusal
/// (<c>SUM(FILTER(A1:A3,A1:A3&gt;100,A:A))</c> is <c>#VALUE!</c> against the oracle's 14 — the same
/// open-range deviation, one slot over). What justifies the refusal rather than merely recording it:
/// <c>ROWS</c> of that same formula is <b>1048576</b> on the oracle, the whole column, so matching it would
/// mean materializing a million-row result from a slot nobody reads.</para>
///
/// <para>Every number above was measured on Aspose.Cells 26.6.0, 2026-09-10, plain and array-entered entry
/// agreeing unless named, and every one has an assertion in <c>FilterTests</c> or <c>DynamicArrayTests</c> —
/// the open-range <c>if_empty</c> row did NOT until the phase's final review found this sentence claiming it,
/// which is the project's most-repeated defect and the reason the claim now reads as a checkable one.</para>
/// </remarks>
[MemoryPackable]
public sealed partial record Filter(Expression[] Arguments) : Function, IArrayProducer
{
    public override ComputedValue Evaluate(EvaluationContext context) =>
        ArrayEvaluation.FirstElement(this, context);

    (bool Succeeds, bool IsArray) IArrayProducer.ProbeArray(EvaluationContext context) =>
        SelectionProducers.ProbeSource(Arguments[0], context)
        && ArrayEvaluation.Probe(Arguments[1], context).Succeeds
            ? (true, true)
            : (false, false);

    bool IArrayProducer.TryBuildArrayOperand(EvaluationContext context, out ArrayOperand operand)
    {
        if (
            !SelectionProducers.TryBuildSource(Arguments[0], context, out var source)
            || !ArrayEvaluation.TryBuildOperand(Arguments[1], context, out var include)
        )
        {
            operand = null!;
            return false;
        }

        operand = Build(source, include, context);
        return true;
    }

    private ArrayOperand Build(ArrayOperand source, ArrayOperand include, EvaluationContext context)
    {
        var (includeRows, includeColumns) = include.IsArray
            ? (include.Rows, include.Columns)
            : (1, 1);

        ArrayAxis axis;

        if (includeRows == source.Rows && includeColumns == 1)
        {
            axis = ArrayAxis.Rows;
        }
        else if (includeColumns == source.Columns && includeRows == 1)
        {
            axis = ArrayAxis.Columns;
        }
        else
        {
            return SelectionProducers.Singleton(Error.Value);
        }

        var length = includeRows * includeColumns;

        // The one-element include: unconvertible text is an error here, and nothing kept is #VALUE!.
        if (length == 1)
        {
            var only = include.At(0, includeRows, includeColumns);

            if (only.CoerceToBoolAllowingTextWords(out var keepAll) is { } error)
            {
                return SelectionProducers.Singleton(error);
            }

            return keepAll
                ? new AxisSelectionOperand(
                    source,
                    axis,
                    SelectionProducers.Everything(source, axis)
                )
                : IfEmpty(context, Error.Value);
        }

        var kept = new List<int>(length);

        for (var i = 0; i < length; i++)
        {
            var element = include.At(i, includeRows, includeColumns);

            if (element.TryGetError(out var error))
            {
                return SelectionProducers.Singleton(error);
            }

            if (element.CoerceToBoolAllowingTextWords(out var keep) is null && keep)
            {
                kept.Add(i);
            }
        }

        return kept.Count == 0
            ? IfEmpty(context, Error.Calc)
            : new AxisSelectionOperand(source, axis, kept.ToArray());
    }

    // The empty result: if_empty when given (built only now — an error there is never read otherwise),
    // else the error the empty case carries. A REFUSED if_empty cannot propagate as a refusal (the probe
    // never looked at it), so it is a loud 1x1 #VALUE! instead.
    private ArrayOperand IfEmpty(EvaluationContext context, Error whenOmitted)
    {
        if (Arguments.Length < 3)
        {
            return SelectionProducers.Singleton(whenOmitted);
        }

        if (!ArrayEvaluation.TryBuildOperand(Arguments[2], context, out var ifEmpty))
        {
            return SelectionProducers.Singleton(Error.Value);
        }

        return ifEmpty.IsArray ? ifEmpty : new SingletonArrayOperand(ifEmpty.Scalar);
    }
}
