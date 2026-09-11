using MemoryPack;

namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// <c>UNIQUE(array, [by_col], [exactly_once])</c> — the distinct rows (or, with <c>by_col</c>, the distinct
/// columns) of <c>array</c> in order of FIRST appearance, as a mini-CSE producer (<see cref="IArrayProducer"/>):
/// <c>UNIQUE(Q1:Q4)</c> over 9, 5, 9, 0 is 9, 5, 0 (<c>SUM</c> 14, <c>COUNTA</c> 3), and a cell shows the
/// top-left (<see cref="ArrayEvaluation.FirstElement"/>). The KEYS are read once at build time; the values
/// stay on demand through <see cref="AxisSelectionOperand"/>, so a kept blank arrives as a blank.
/// </summary>
/// <remarks>
/// <para><b>Keys.</b> Two candidates (a row, or a column) are the same when every cell pair is the same
/// KIND and the same value: numbers numerically, booleans by value, text ORDINALLY — case-SENSITIVE, as
/// measured: <c>UNIQUE(C1:C3)</c> over "a", "A", "b" keeps all three with "A" second, and rows (1,"a"),
/// (1,"a"), (1,"A") are two — while <c>COUNTIF(C1:C3,"a")</c> = 2 and <c>MATCH("A",UNIQUE(C1:C3),0)</c> = 1
/// stay case-insensitive beside it. Microsoft's UNIQUE page as fetched 2026-09-10 says nothing about case
/// either way, so the measurement stands alone (P0). Kinds never cross: 1, <c>"1"</c>, TRUE, 1 are three
/// rows. A BLANK is its own key, equal to neither 0 nor <c>""</c> nor FALSE — 0, blank, <c>""</c>, FALSE are
/// four rows, and <c>UNIQUE(A5:A8)</c> over 7, blank, "t", 7 is three with <c>ISBLANK</c> of the second
/// TRUE, <c>COUNTA</c> 2, <c>COUNT</c> 1 — which is why <see cref="ValueCoercion.AreEqual"/> (blank equals
/// the empty of the other side, non-transitively) is not used and no normalization happens. An ERROR is a
/// key like any other, equal to the same error code: 2, <c>#DIV/0!</c>, 3, <c>#N/A</c>, <c>#DIV/0!</c> keeps
/// four rows with the first <c>#DIV/0!</c> second and <c>#N/A</c> fourth; <c>UNIQUE(E1:E3)</c> keeps its
/// error row (<c>ROWS</c> 3, <c>INDEX(…,2)</c> = <c>#DIV/0!</c>).</para>
///
/// <para><b>Flags</b> coerce as in <see cref="SelectionProducers"/>, errors in argument order.
/// <c>by_col</c> compares columns: <c>COLUMNS(UNIQUE(A1:B3,TRUE))</c> = 2, <c>UNIQUE(A1:A3,TRUE)</c> is the
/// one 3x1 column (<c>SUM</c> 14). <c>exactly_once</c> keeps the candidates that occur ONCE — 9, 5, 9, 0
/// gives 5, 0 — following Microsoft's page ("all distinct rows or columns that occur exactly once"), and NOT
/// the oracle. The oracle keeps the DISTINCT-count shape and OVERWRITES ITS FRONT with the once-occurring
/// values, leaving every distinct element the once-list does not reach exactly where it already was. The
/// clearest fixture is 1, 1, 2, 3, 3, 4: the distinct rows are 1, 2, 3, 4, only 2 and 4 occur once, and the
/// oracle answers <b>2, 4, 3, 4</b> — a four-row result carrying a value that occurs twice (3) and a
/// duplicate of one that does not (4), so it contradicts its own row count in two ways at once. That is why
/// it is not matched. The same overwrite is why the oracle answers a 1-row 4 for
/// <c>UNIQUE(S1:S2,FALSE,TRUE)</c> over 4, 4, where nothing occurs once and so nothing is overwritten —
/// here that is the empty result, a 1x1 <c>#CALC!</c> (<see cref="ArrayShaping"/>: never a 0-extent array).
/// An earlier version of this paragraph said the oracle "pads by repeating the last kept value". That model
/// coincides on 9, 5, 9, 0 and on 1, 2, 2, 3 and is wrong everywhere the distinct tail differs from the last
/// kept value: it predicts 5, 9, 9 for 5, 9, 0, 0 (measured 5, 9, 0) and 8, 8, 8 for 7, 7, 8, 9, 9
/// (measured 8, 8, 9). All measured on Aspose.Cells 26.6.0, 2026-09-10, both entry modes.</para>
///
/// <para><b>Refusal.</b> The build returns <c>false</c> exactly when the array's own build is refused (an
/// open range) and <c>ProbeArray</c> asks the same question; every other failure is the producer's own 1x1
/// <see cref="SingletonArrayOperand"/>. A 1x1 source is a 1x1 result (<c>SUM(UNIQUE(A1))</c> = 5,
/// <c>=UNIQUE("x")</c> = "x", <c>SUM(UNIQUE(1/0))</c> = <c>#DIV/0!</c>).</para>
///
/// <para>Every number above except the <c>exactly_once</c> shape was measured on Aspose.Cells 26.6.0,
/// 2026-09-10, plain and array-entered entry agreeing; pinned in <c>UniqueTests</c> and
/// <c>DynamicArrayTests</c>.</para>
/// </remarks>
[MemoryPackable]
public sealed partial record Unique(Expression[] Arguments) : Function, IArrayProducer
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
        if (
            SelectionProducers.ReadFlag(Arguments, 1, context, out var byColumn) is
            { } byColumnError
        )
        {
            return SelectionProducers.Singleton(byColumnError);
        }

        if (
            SelectionProducers.ReadFlag(Arguments, 2, context, out var exactlyOnce) is
            { } exactlyOnceError
        )
        {
            return SelectionProducers.Singleton(exactlyOnceError);
        }

        var axis = byColumn ? ArrayAxis.Columns : ArrayAxis.Rows;
        var length = SelectionProducers.Extent(source, axis);
        var width = SelectionProducers.CrossExtent(source, axis);

        // Group positions by key in first-appearance order: the group's first position and its size.
        var groups = new Dictionary<ComputedValue[], int>(KeyEquality.Instance);
        var firstPositions = new List<int>();
        var sizes = new List<int>();

        for (var position = 0; position < length; position++)
        {
            var key = new ComputedValue[width];

            for (var offset = 0; offset < width; offset++)
            {
                key[offset] = SelectionProducers.ElementAt(source, axis, position, offset);
            }

            if (groups.TryGetValue(key, out var group))
            {
                sizes[group]++;
            }
            else
            {
                groups.Add(key, groups.Count);
                firstPositions.Add(position);
                sizes.Add(1);
            }
        }

        var kept = new List<int>(firstPositions.Count);

        for (var group = 0; group < firstPositions.Count; group++)
        {
            if (!exactlyOnce || sizes[group] == 1)
            {
                kept.Add(firstPositions[group]);
            }
        }

        return kept.Count == 0
            ? SelectionProducers.Singleton(Error.Calc)
            : new AxisSelectionOperand(source, axis, kept.ToArray());
    }

    // Key equality as the type remarks state it: same kind, same value, text ordinal, blank its own value,
    // an error its code. The hash agrees with it (a zero hashes as one value whichever its sign).
    private sealed class KeyEquality : IEqualityComparer<ComputedValue[]>
    {
        public static readonly KeyEquality Instance = new();

        public bool Equals(ComputedValue[]? left, ComputedValue[]? right)
        {
            if (left is null || right is null || left.Length != right.Length)
            {
                return left is null && right is null;
            }

            for (var i = 0; i < left.Length; i++)
            {
                if (!SameKey(left[i], right[i]))
                {
                    return false;
                }
            }

            return true;
        }

        public int GetHashCode(ComputedValue[] key)
        {
            var hash = new HashCode();

            foreach (var cell in key)
            {
                hash.Add(cell.Kind);
                hash.Add(KeyHash(cell));
            }

            return hash.ToHashCode();
        }

        private static bool SameKey(in ComputedValue left, in ComputedValue right)
        {
            if (left.Kind != right.Kind)
            {
                return false;
            }

            switch (left.Kind)
            {
                case ComputedValueKind.Blank:
                    return true;

                case ComputedValueKind.Number:
                    left.TryGetNumber(out var leftNumber);
                    right.TryGetNumber(out var rightNumber);
                    return leftNumber == rightNumber;

                case ComputedValueKind.Boolean:
                    left.TryGetBoolean(out var leftBoolean);
                    right.TryGetBoolean(out var rightBoolean);
                    return leftBoolean == rightBoolean;

                case ComputedValueKind.Text:
                    left.TryGetText(out var leftText);
                    right.TryGetText(out var rightText);
                    return string.Equals(leftText, rightText, StringComparison.Ordinal);

                case ComputedValueKind.Error:
                    left.TryGetError(out var leftError);
                    right.TryGetError(out var rightError);
                    return leftError.Equals(rightError);

                default:
                    return false;
            }
        }

        private static int KeyHash(in ComputedValue cell)
        {
            switch (cell.Kind)
            {
                case ComputedValueKind.Number:
                    cell.TryGetNumber(out var number);
                    return number == 0 ? 0 : number.GetHashCode();

                case ComputedValueKind.Boolean:
                    cell.TryGetBoolean(out var boolean);
                    return boolean ? 1 : 0;

                case ComputedValueKind.Text:
                    cell.TryGetText(out var text);
                    return string.GetHashCode(text, StringComparison.Ordinal);

                case ComputedValueKind.Error:
                    cell.TryGetError(out var error);
                    return error.GetHashCode();

                default:
                    return 0;
            }
        }
    }
}
