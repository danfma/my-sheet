namespace Danfma.MySheet.Expressions.Mathematics;

/// <summary>
/// The aggregate core shared by <see cref="Subtotal"/> and AGGREGATE: the per-cell scan that skips nested
/// aggregate cells, the accumulator those two feed, and the two halves of the function_num map (the 1-13
/// fold over the whole population and the 14-19 positional selection).
///
/// <para>It lives here, <c>internal</c> in its own file rather than the codebase's local
/// <c>file static class</c> idiom (<c>SumOfPairs</c>), because it has TWO consumers and a file-local type is
/// invisible across files; a second copy of the 140-line scan — five reference shapes, each carrying a
/// measured invariant — would let SUBTOTAL and AGGREGATE drift. <c>Mathematics</c> is the right namespace
/// because <see cref="IsNested"/> needs both node types by name.</para>
/// </summary>
internal static class AggregateCodes
{
    /// <summary>Which nested aggregate cells a scan drops from the population. Microsoft documents the rule
    /// for SUBTOTAL as "nested subtotals are ignored" and for AGGREGATE (options 0-3) as nested SUBTOTAL
    /// *and* AGGREGATE, so the predicate is a parameter, not a constant.</summary>
    internal enum NestedSkip : byte
    {
        None,
        Subtotal,
        SubtotalAndAggregate,
    }

    /// <summary>
    /// THE entry point of both callers: routes one argument into <paramref name="accumulator"/>.
    ///
    /// <para>A non-<see cref="Reference"/> argument that the mini-CSE can evaluate element-wise — the
    /// <c>ROW(A1:A3)</c> / <c>(A1:A3&lt;&gt;0)*1</c> shapes — is STREAMED, exactly as SUM/COUNT/AVERAGE do
    /// through <c>NumericAggregation.Fold</c>'s default arm, so SUBTOTAL folds a computed array the way Excel
    /// does (<c>=SUBTOTAL(9,ROW(A1:A3))</c> is 6, like <c>=SUM(ROW(A1:A3))</c>). Reaching such an argument
    /// through <see cref="Gather"/>'s <c>default:</c> instead would evaluate it as a SCALAR — silently the
    /// first row number for <c>ROW(range)</c>, and <c>#VALUE!</c> for an operation over a range, which has no
    /// scalar value. Every <see cref="Reference"/> shape keeps the cell-by-cell scan below, which is what
    /// carries the nested-aggregate skip.</para>
    /// </summary>
    public static Error? Feed(
        Expression argument,
        EvaluationContext context,
        ref Accumulator accumulator,
        NestedSkip skip
    ) =>
        ArrayEvaluation.TryStream(argument, context, out var stream)
            ? CollectStream(stream, ref accumulator)
            : Gather(argument, context, ref accumulator, skip);

    // Walks a ref argument cell by cell — numerically, wherever the shape allows it, so a big range pays
    // neither a CellAddress.ToId() build (to ask "is this cell's stored formula a nested aggregate") nor a
    // second one (to read its value) per cell — folding straight into the accumulator instead of building an
    // intermediate List<ComputedValue> that the fold used to re-walk to extract doubles. Returns the Error to
    // propagate (a numeric aggregate short-circuits on the first cell error, exactly as the old two-pass
    // Aggregate did in scan order); COUNT/COUNTA never produce one, matching their standalone functions.
    public static Error? Gather(
        Expression argument,
        EvaluationContext context,
        ref Accumulator accumulator,
        NestedSkip skip
    )
    {
        switch (argument)
        {
            case RangeReference range:
            {
                var workbook = context.Workbook;
                var sheet = workbook.Sheets[range.SheetName];
                var bounds = range.GetBounds();
                var handle = workbook.ResolveDenseHandle(range.SheetName);

                for (var column = bounds.LeftColumn; column <= bounds.RightColumn; column++)
                {
                    for (var row = bounds.TopRow; row <= bounds.BottomRow; row++)
                    {
                        // `skip != None` first: options 4-7 answer the nested question with a constant, so
                        // they must not pay a per-cell expression lookup to reach it.
                        if (
                            skip != NestedSkip.None
                            && sheet.TryGetCellExpressionDense(column, row, out var expression)
                            && IsNested(expression, skip)
                        )
                        {
                            continue;
                        }

                        var value = workbook.GetCellValueDense(
                            handle,
                            range.SheetName,
                            column,
                            row
                        );

                        if (accumulator.Add(value) is { } error)
                        {
                            return error;
                        }
                    }
                }

                return null;
            }

            case OpenRangeReference open:
            {
                var workbook = context.Workbook;
                var sheet = workbook.Sheets[open.SheetName];
                var handle = workbook.ResolveDenseHandle(open.SheetName);

                // The index-backed (column, row) walk — no per-cell id string at all, unlike PopulatedIds.
                foreach (var (column, row) in open.PopulatedCells(context))
                {
                    if (
                        skip != NestedSkip.None
                        && sheet.TryGetCellExpressionDense(column, row, out var expression)
                        && IsNested(expression, skip)
                    )
                    {
                        continue;
                    }

                    var value = workbook.GetCellValueDense(handle, open.SheetName, column, row);

                    if (accumulator.Add(value) is { } error)
                    {
                        return error;
                    }
                }

                return null;
            }

            case CellReference cell:
            {
                var workbook = context.Workbook;
                var sheet = workbook.Sheets[cell.SheetName];

                // Every CellReference the parser produces carries a canonical A1 id, so this hits in
                // practice; a non-canonical id (a defined-name/host edge case) falls back to the exact
                // original string path below.
                if (CellAddress.TryGetColumnRow(cell.Id, out var column, out var row))
                {
                    if (
                        skip != NestedSkip.None
                        && sheet.TryGetCellExpressionDense(column, row, out var expression)
                        && IsNested(expression, skip)
                    )
                    {
                        return null;
                    }

                    var handle = workbook.ResolveDenseHandle(cell.SheetName);
                    var value = workbook.GetCellValueDense(handle, cell.SheetName, column, row);

                    return accumulator.Add(value);
                }

                // Same hoist as the dense probes above: options 4-7 must not pay the sheet lookup to reach
                // a constant answer.
                if (skip == NestedSkip.None || !IsNested(sheet[cell.Id], skip))
                {
                    return accumulator.Add(cell.Evaluate(context));
                }

                return null;
            }

            case UnionReference union:
                foreach (var area in union.Areas)
                {
                    if (Gather(area, context, ref accumulator, skip) is { } error)
                    {
                        return error;
                    }
                }

                return null;

            // Phase 2 audit (shared-formula delta production): a SUBTOTAL ref argument written INSIDE a
            // shared-formula master is an anchored node, not a plain CellReference/RangeReference — it fell
            // to the `default` branch below before this fix, which evaluates the argument directly and
            // therefore skips BOTH the nested-subtotal exclusion rule AND (for a range) even a correct value
            // (AnchoredRangeReference.Evaluate always returns #VALUE!, like RangeReference, since a range has
            // no scalar value). Resolving to the concrete, delta-applied twin and re-dispatching through this
            // same switch reuses the RangeReference/CellReference cases above exactly — no logic duplicated.
            case AnchoredCellReference
            or AnchoredRangeReference:
            {
                ((Reference)argument).TryResolveReference(context, out var resolved);
                return Gather(resolved!, context, ref accumulator, skip);
            }

            default:
                var computed = argument.Evaluate(context);

                // A reference produced by a function (OFFSET, CHOOSE, …) still carries the actual
                // Reference node, so the nested-aggregate skip applies to it as well.
                return computed.TryGetReference(out var reference)
                    ? Gather(reference, context, ref accumulator, skip)
                    : accumulator.Add(computed);
        }
    }

    /// <summary>Feeds a mini-CSE array element by element into the SAME accumulator, so every option rule —
    /// ignoreErrors, COUNTA's error exclusion — stays in one place. Used for an array-eligible argument of
    /// any code, and by AGGREGATE's array form for 16-19, which need the whole population anyway (the bounded
    /// heap of <c>OrderSelection.KthValueStreaming</c> buys nothing there).</summary>
    public static Error? CollectStream(
        ArrayEvaluation.ArrayStream stream,
        ref Accumulator accumulator
    )
    {
        foreach (var element in stream)
        {
            if (accumulator.Add(element) is { } error)
            {
                return error;
            }
        }

        return null;
    }

    // The two rules are deliberately DIFFERENT widths: SUBTOTAL's page documents only "nested subtotals are
    // ignored", while "nested SUBTOTAL and AGGREGATE functions" is AGGREGATE's own options-table wording
    // (options 0-3). The converse — SUBTOTAL skipping a nested AGGREGATE — is unverified and is not guessed
    // into the engine, which is why NestedSkip.Subtotal stays narrow.
    private static bool IsNested(Expression? expression, NestedSkip skip) =>
        skip switch
        {
            NestedSkip.None => false,
            NestedSkip.Subtotal => expression is Subtotal,
            _ => expression is Subtotal or Aggregate,
        };

    // Accumulates the exact shape the aggregate needs, directly from the per-cell scan — no intermediate
    // List<ComputedValue> the way the old Aggregate() re-walked to build its List<double>. COUNT (2) and
    // COUNTA (3) never propagate cell errors, like their standalone functions, so they only ever tally
    // counts; every other code needs the numeric population (STDEV/VAR/AVERAGE all fold over the whole list)
    // and DOES propagate the first error it meets, matching the old scan-order short-circuit.
    //
    // ignoreErrors is AGGREGATE's option bit 1. The accumulator ALREADY excludes error cells from _numbers
    // (the default arm returns before _numbers.Add), so for codes 1-13 "ignore errors" is purely the
    // suppression of that propagation channel — plus one genuine behavioural difference in COUNTA, which
    // otherwise counts the very error cells the option says to ignore (measured: =SUBTOTAL(3,A1:A3) with an
    // error cell → 3, so AGGREGATE(3,6,A1:A3) must be 2). COUNT needs no rule: it only tallies Numbers.
    internal struct Accumulator(int code, bool ignoreErrors)
    {
        private int _count; // code 2 (COUNT): cells whose value is a Number
        private int _countA; // code 3 (COUNTA): cells whose value is not Blank
        private readonly List<double> _numbers = code is 2 or 3 ? null! : new List<double>();

        /// <summary>The numeric population gathered so far — what AGGREGATE's positional codes (14-19) select
        /// from instead of calling <see cref="Finish"/>. Null for codes 2 and 3, which keep only a tally.</summary>
        public readonly List<double> Numbers => _numbers;

        public Error? Add(ComputedValue value)
        {
            switch (code)
            {
                case 2:
                    if (value.Kind == ComputedValueKind.Number)
                    {
                        _count++;
                    }

                    return null;

                case 3:
                    if (ignoreErrors && value.Kind == ComputedValueKind.Error)
                    {
                        return null;
                    }

                    if (value.Kind != ComputedValueKind.Blank)
                    {
                        _countA++;
                    }

                    return null;

                default:
                    if (value.TryGetError(out var error))
                    {
                        return ignoreErrors ? null : error;
                    }

                    // Referenced semantics: text, logicals and blanks are ignored.
                    if (value.TryGetNumber(out var number))
                    {
                        _numbers.Add(number);
                    }

                    return null;
            }
        }

        public readonly ComputedValue Finish()
        {
            if (code == 2)
            {
                return ComputedValue.Number(_count);
            }

            if (code == 3)
            {
                return ComputedValue.Number(_countA);
            }

            return Fold(code, _numbers);
        }
    }

    /// <summary>The 1-13 half of the function_num map: a fold over the whole numeric population. Codes 2 and
    /// 3 never reach it — the accumulator answers them from its tallies.
    ///
    /// <para>MUTATES <paramref name="numbers"/>: code 12 (MEDIAN) SORTS it ascending in place, so a caller
    /// that keeps the list — to feed <see cref="Positional"/>, which requires ascending order — must not
    /// depend on the original scan order afterwards. Every other code reads the list without reordering it,
    /// and code 13 (MODE.SNGL) positively REQUIRES the scan order for its tie-break.</para></summary>
    public static ComputedValue Fold(int code, List<double> numbers)
    {
        switch (code)
        {
            case 1: // AVERAGE
                return numbers.Count == 0
                    ? ComputedValue.Error(Error.DivZero)
                    : ComputedValue.Number(StatisticsMath.Mean(numbers));

            case 4: // MAX
            case 5: // MIN
            {
                if (numbers.Count == 0)
                {
                    return ComputedValue.Number(0);
                }

                var extreme = numbers[0];

                foreach (var number in numbers)
                {
                    if (code == 4 ? number > extreme : number < extreme)
                    {
                        extreme = number;
                    }
                }

                return ComputedValue.Number(extreme);
            }

            case 6: // PRODUCT
            {
                if (numbers.Count == 0)
                {
                    return ComputedValue.Number(0);
                }

                var product = 1.0;

                foreach (var number in numbers)
                {
                    product *= number;
                }

                return ComputedValue.Number(product);
            }

            case 7: // STDEV.S
                return StatisticsMath.SampleVariance(numbers, out var sampleSd) is { } sdError
                    ? ComputedValue.Error(sdError)
                    : ComputedValue.Number(Math.Sqrt(sampleSd));

            case 8: // STDEV.P
                return
                    StatisticsMath.PopulationVariance(numbers, out var populationSd) is { } sdpError
                    ? ComputedValue.Error(sdpError)
                    : ComputedValue.Number(Math.Sqrt(populationSd));

            case 9: // SUM
            {
                var total = 0.0;

                foreach (var number in numbers)
                {
                    total += number;
                }

                return ComputedValue.Number(total);
            }

            case 10: // VAR.S
                return
                    StatisticsMath.SampleVariance(numbers, out var sampleVariance) is { } varError
                    ? ComputedValue.Error(varError)
                    : ComputedValue.Number(sampleVariance);

            case 11: // VAR.P
                return
                    StatisticsMath.PopulationVariance(numbers, out var populationVariance)
                        is { } varpError
                    ? ComputedValue.Error(varpError)
                    : ComputedValue.Number(populationVariance);

            case 12: // MEDIAN — the fold takes an ASCENDING population, so the scan order is sorted away here
                numbers.Sort();

                return StatisticsMath.Median(numbers, out var median) is { } medianError
                    ? ComputedValue.Error(medianError)
                    : ComputedValue.Number(median);

            case 13: // MODE.SNGL — deliberately NOT sorted: the documented tie-break is the FIRST value in
                // scan order, so sorting would change which of two equally frequent values wins.
                return StatisticsMath.Mode(numbers, out var mode) is { } modeError
                    ? ComputedValue.Error(modeError)
                    : ComputedValue.Number(mode);

            default:
                // Unreachable: SUBTOTAL rejects a function_num outside 1-11 and AGGREGATE outside 1-19 (and
                // routes 14-19 to Positional) before getting here. A real guard rather than a catch-all arm,
                // so a future code can never fall silently into the last fold.
                return ComputedValue.Error(Error.Value);
        }
    }

    /// <summary>The 14-19 half of the function_num map: a positional selection over the ASCENDING-sorted
    /// population plus AGGREGATE's fourth argument (k / quart).
    ///
    /// <para><paramref name="kArgument"/> is coerced AFTER the population is gathered, preserving the order
    /// <c>OrderSelection.SortedArrayAndScalar</c> documents — the array's first error precedes the scalar's —
    /// which is observable under options 0/1/4/5, where errors still propagate.</para></summary>
    public static ComputedValue Positional(
        int code,
        IReadOnlyList<double> sorted,
        Expression kArgument,
        EvaluationContext context
    )
    {
        if (kArgument.Evaluate(context).CoerceToNumber(out var k) is { } kError)
        {
            return ComputedValue.Error(kError);
        }

        switch (code)
        {
            case 14: // LARGE
                return OrderSelection.KthOfSorted(sorted, k, largest: true);

            case 15: // SMALL
                return OrderSelection.KthOfSorted(sorted, k, largest: false);

            case 16: // PERCENTILE.INC
                return
                    StatisticsMath.PercentileInclusive(sorted, k, out var percentileInclusive)
                        is { } percentileIncError
                    ? ComputedValue.Error(percentileIncError)
                    : ComputedValue.Number(percentileInclusive);

            case 17: // QUARTILE.INC
                return
                    StatisticsMath.QuartileInclusive(sorted, k, out var quartileInclusive)
                        is { } quartileIncError
                    ? ComputedValue.Error(quartileIncError)
                    : ComputedValue.Number(quartileInclusive);

            case 18: // PERCENTILE.EXC
                return
                    StatisticsMath.PercentileExclusive(sorted, k, out var percentileExclusive)
                        is { } percentileExcError
                    ? ComputedValue.Error(percentileExcError)
                    : ComputedValue.Number(percentileExclusive);

            case 19: // QUARTILE.EXC
                return
                    StatisticsMath.QuartileExclusive(sorted, k, out var quartileExclusive)
                        is { } quartileExcError
                    ? ComputedValue.Error(quartileExcError)
                    : ComputedValue.Number(quartileExclusive);

            default:
                // Unreachable: AGGREGATE routes only 14-19 here (1-13 go to Fold). A real guard rather than
                // a catch-all arm, mirroring Fold, so a future code can never fall silently into the last
                // selection.
                return ComputedValue.Error(Error.Value);
        }
    }
}
