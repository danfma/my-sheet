namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// What <see cref="Filter"/>, <see cref="Sort"/> and <see cref="Unique"/> share. The three are ONE operation
/// — a selection or permutation of one axis of a SOURCE array, carried by <see cref="AxisSelectionOperand"/>
/// — and differ only in how the selection is computed at build time, so the source rule, the omitted-argument
/// rule, the flag coercion and the axis arithmetic live here once instead of three times.
/// </summary>
/// <remarks>
/// <para><b>The source.</b> <see cref="ProbeSource"/> and <see cref="TryBuildSource"/> are twins in the
/// sense <see cref="IArrayProducer"/> requires: the build returns <c>false</c> exactly when the probe answers
/// <c>Succeeds = false</c> — an open range somewhere below, the mini-CSE's cost guard — and in no other case.
/// An ARRAY source is used as it is; a SCALAR source becomes a 1x1 <see cref="SingletonArrayOperand"/>,
/// because <see cref="ScalarOperand"/> reports 0x0 and can never be a result (blocker B1:
/// <c>SUM(FILTER(A1,TRUE))</c>, <c>SUM(SORT(A1))</c> and <c>SUM(UNIQUE(A1))</c> are all 5, and a bare
/// <c>=SORT("x")</c> is <c>"x"</c>; Aspose.Cells 26.6.0, 2026-09-10, both entry modes). A scalar that holds a
/// REFERENCE — the value a reference-returning node such as <c>INDIRECT("A1:A3")</c> or
/// <c>OFFSET(A1,0,0,3,1)</c> evaluates to, which the builder treats as an opaque scalar — is NOT wrapped
/// (correction B1's letter): a rectangle resolves to its <c>RangeOperand</c> through
/// <see cref="ArrayEvaluation.BuildRange"/>, so <c>SUM(SORT(INDIRECT("A1:A3")))</c> and
/// <c>SUM(UNIQUE(OFFSET(A1,0,0,3,1)))</c> are 14 as measured (same oracle, both modes); a single cell is its
/// value; any other reference shape (an open range, a union) is a loud 1x1 <c>#VALUE!</c>.</para>
///
/// <para><b>Omitted arguments</b> follow <c>SEQUENCE</c>'s rule: an absent argument or the parser's
/// <see cref="BlankValue"/> for an empty slot is omitted, while a blank CELL is a value (it coerces to 0 as a
/// number and to FALSE as a flag). Measured: <c>INDEX(SORT(A1:A3,,-1),1)</c> = 9 (empty <c>sort_index</c>
/// slot defaults to 1), <c>INDEX(SORT(A1:A3,A9),1)</c> = <c>#VALUE!</c> (blank cell → 0 → out of range),
/// <c>ROWS(UNIQUE(Q1:Q4,))</c> = 3 and <c>ROWS(UNIQUE(Q1:Q4,A9))</c> = 3 (empty slot and blank cell are
/// both FALSE here, so the two rules coincide for a flag). <c>FILTER</c>'s <c>if_empty</c> is the one slot
/// that does NOT follow it — see <see cref="Filter"/>.</para>
///
/// <para><b>Flags</b> (<c>by_col</c>, <c>exactly_once</c>) coerce with the text words <c>TRUE</c>/<c>FALSE</c>
/// accepted, exactly as <c>FILTER</c>'s include does: <c>ROWS(UNIQUE(Q1:Q4,"TRUE"))</c> = 4 (by column),
/// <c>ROWS(UNIQUE(Q1:Q4,2))</c> = 4, <c>INDEX(SORT(A1:B3,1,1,"TRUE"),1,1)</c> = 1, while <c>"x"</c> is
/// <c>#VALUE!</c> and an error propagates, in argument order (<c>ROWS(UNIQUE(Q1:Q4,1/0,"x"))</c> =
/// <c>#DIV/0!</c>, <c>ROWS(UNIQUE(Q1:Q4,"x",1/0))</c> = <c>#VALUE!</c>). Same oracle, both modes.</para>
/// </remarks>
internal static class SelectionProducers
{
    /// <summary>The shape twin of <see cref="TryBuildSource"/>: whether the source's own build succeeds.</summary>
    public static bool ProbeSource(Expression expression, EvaluationContext context) =>
        ArrayEvaluation.Probe(expression, context).Succeeds;

    /// <summary>
    /// Builds a producer's source as an ARRAY operand (see the type remarks), or returns <c>false</c> when
    /// the source's build is refused.
    /// </summary>
    public static bool TryBuildSource(
        Expression expression,
        EvaluationContext context,
        out ArrayOperand source
    )
    {
        if (!ArrayEvaluation.TryBuildOperand(expression, context, out var built))
        {
            source = null!;
            return false;
        }

        if (built.IsArray)
        {
            source = built;
            return true;
        }

        var scalar = built.Scalar;

        source = scalar.TryGetReference(out var reference)
            ? reference switch
            {
                RangeReference range => ArrayEvaluation.BuildRange(range, context),
                CellReference cell => new SingletonArrayOperand(cell.Evaluate(context)),
                _ => Singleton(Error.Value),
            }
            : new SingletonArrayOperand(scalar);
        return true;
    }

    /// <summary>An absent argument, or the parser's <see cref="BlankValue"/> for an empty slot.</summary>
    public static bool IsOmitted(Expression[] arguments, int index) =>
        index >= arguments.Length || arguments[index] is BlankValue;

    /// <summary>
    /// A boolean flag: FALSE when omitted, otherwise the argument evaluated ONCE and coerced with the text
    /// words accepted (<see cref="ValueCoercion.CoerceToBoolAllowingTextWords"/>). Returns the error to
    /// propagate, or <c>null</c>.
    /// </summary>
    public static Error? ReadFlag(
        Expression[] arguments,
        int index,
        EvaluationContext context,
        out bool flag
    )
    {
        if (IsOmitted(arguments, index))
        {
            flag = false;
            return null;
        }

        return arguments[index].Evaluate(context).CoerceToBoolAllowingTextWords(out flag);
    }

    /// <summary>The source's extent ALONG <paramref name="axis"/> — the number of candidates to select from.</summary>
    public static int Extent(ArrayOperand source, ArrayAxis axis) =>
        axis is ArrayAxis.Rows ? source.Rows : source.Columns;

    /// <summary>The source's extent ACROSS <paramref name="axis"/> — the width of one candidate.</summary>
    public static int CrossExtent(ArrayOperand source, ArrayAxis axis) =>
        axis is ArrayAxis.Rows ? source.Columns : source.Rows;

    /// <summary>
    /// The element at <paramref name="position"/> along <paramref name="axis"/> and <paramref name="offset"/>
    /// across it, read at the source's OWN extent (so a composite source is read exactly as anywhere else).
    /// </summary>
    public static ComputedValue ElementAt(
        ArrayOperand source,
        ArrayAxis axis,
        int position,
        int offset
    )
    {
        var (row, column) = axis is ArrayAxis.Rows ? (position, offset) : (offset, position);

        return source.At(row * source.Columns + column, source.Rows, source.Columns);
    }

    /// <summary>The identity selection: every position along <paramref name="axis"/>, in order.</summary>
    public static int[] Everything(ArrayOperand source, ArrayAxis axis)
    {
        var selection = new int[Extent(source, axis)];

        for (var i = 0; i < selection.Length; i++)
        {
            selection[i] = i;
        }

        return selection;
    }

    public static SingletonArrayOperand Singleton(Error error) => new(ComputedValue.Error(error));
}
