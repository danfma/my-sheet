using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record Subtotal(Expression[] Arguments) : Function
{
    // SUBTOTAL(function_num, ref1, [ref2], …) — applies the aggregate selected by function_num
    // (1-11; 101-111 behave identically here: MySheet has no hidden-row model, a documented model
    // limit) while IGNORING any referenced cell whose own formula is a SUBTOTAL node, so stacked
    // subtotal rows are not double counted (the real nested-subtotal rule). Rows excluded by a
    // filter do not apply either (no filter model). An invalid function_num → #VALUE!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var rawCode) is { } codeError)
        {
            return ComputedValue.Error(codeError);
        }

        var code = (int)Math.Truncate(rawCode);

        if (code is >= 101 and <= 111)
        {
            code -= 100; // "ignore hidden rows" variants: same behaviour without a hidden-row model
        }

        if (code is < 1 or > 11)
        {
            return ComputedValue.Error(Error.Value);
        }

        // A reference argument to a missing sheet is a structural #REF! that short-circuits SUBTOTAL, before
        // any per-cell scan (which would otherwise index a non-existent sheet, or be swallowed by the
        // COUNT/COUNTA codes).
        if (ReferenceGuard.MissingSheet(Arguments[1..], context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var accumulator = new SubtotalAccumulator(code);

        foreach (var argument in Arguments[1..])
        {
            if (GatherSkippingSubtotals(argument, context, ref accumulator) is { } error)
            {
                return ComputedValue.Error(error);
            }
        }

        return accumulator.Finish();
    }

    // Walks a ref argument cell by cell — numerically, wherever the shape allows it, so a big range pays
    // neither a CellAddress.ToId() build (to ask "is this cell's stored formula a SUBTOTAL") nor a second one
    // (to read its value) per cell — folding straight into the accumulator instead of building an intermediate
    // List<ComputedValue> that Aggregate used to re-walk to extract doubles. Returns the Error to propagate
    // (a numeric aggregate short-circuits on the first cell error, exactly as the old two-pass Aggregate did
    // in scan order); COUNT/COUNTA never produce one, matching their standalone functions.
    private static Error? GatherSkippingSubtotals(
        Expression argument,
        EvaluationContext context,
        ref SubtotalAccumulator accumulator
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
                        if (
                            sheet.TryGetCellExpressionDense(column, row, out var expression)
                            && expression is Subtotal
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
                        sheet.TryGetCellExpressionDense(column, row, out var expression)
                        && expression is Subtotal
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
                        sheet.TryGetCellExpressionDense(column, row, out var expression)
                        && expression is Subtotal
                    )
                    {
                        return null;
                    }

                    var handle = workbook.ResolveDenseHandle(cell.SheetName);
                    var value = workbook.GetCellValueDense(handle, cell.SheetName, column, row);

                    return accumulator.Add(value);
                }

                if (sheet[cell.Id] is not Subtotal)
                {
                    return accumulator.Add(cell.Evaluate(context));
                }

                return null;
            }

            case UnionReference union:
                foreach (var area in union.Areas)
                {
                    if (GatherSkippingSubtotals(area, context, ref accumulator) is { } error)
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
                return GatherSkippingSubtotals(resolved!, context, ref accumulator);
            }

            default:
                var computed = argument.Evaluate(context);

                // A reference produced by a function (OFFSET, CHOOSE, …) still carries the actual
                // Reference node, so the nested-subtotal skip applies to it as well.
                return computed.TryGetReference(out var reference)
                    ? GatherSkippingSubtotals(reference, context, ref accumulator)
                    : accumulator.Add(computed);
        }
    }

    // Accumulates the exact shape SUBTOTAL's aggregate needs, directly from the per-cell scan — no
    // intermediate List<ComputedValue> the way the old Aggregate() re-walked to build its List<double>.
    // COUNT (2) and COUNTA (3) never propagate cell errors, like their standalone functions, so they only
    // ever tally counts; every other code needs the numeric population (STDEV/VAR/AVERAGE all fold over the
    // whole list) and DOES propagate the first error it meets, matching the old scan-order short-circuit.
    private struct SubtotalAccumulator(int code)
    {
        private int _count; // code 2 (COUNT): cells whose value is a Number
        private int _countA; // code 3 (COUNTA): cells whose value is not Blank
        private readonly List<double> _numbers = code is 2 or 3 ? null! : new List<double>();

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
                    if (value.Kind != ComputedValueKind.Blank)
                    {
                        _countA++;
                    }

                    return null;

                default:
                    if (value.TryGetError(out var error))
                    {
                        return error;
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

            return Aggregate(code, _numbers);
        }
    }

    private static ComputedValue Aggregate(int code, List<double> numbers)
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

            default: // 11: VAR.P
                return
                    StatisticsMath.PopulationVariance(numbers, out var populationVariance)
                        is { } varpError
                    ? ComputedValue.Error(varpError)
                    : ComputedValue.Number(populationVariance);
        }
    }
}
