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

        var values = new List<ComputedValue>();

        foreach (var argument in Arguments[1..])
        {
            GatherSkippingSubtotals(argument, context, values);
        }

        return Aggregate(code, values);
    }

    // Walks a ref argument cell by cell so each cell's STORED EXPRESSION can be inspected: a cell
    // whose formula is a SUBTOTAL call is skipped entirely.
    private static void GatherSkippingSubtotals(
        Expression argument,
        EvaluationContext context,
        List<ComputedValue> values
    )
    {
        switch (argument)
        {
            case RangeReference range:
            {
                var sheet = context.Workbook.Sheets[range.SheetName];
                var start = CellAddress.Parse(range.StartId);
                var end = CellAddress.Parse(range.EndId);
                var minColumn = Math.Min(start.Column, end.Column);
                var maxColumn = Math.Max(start.Column, end.Column);
                var minRow = Math.Min(start.Row, end.Row);
                var maxRow = Math.Max(start.Row, end.Row);

                for (var column = minColumn; column <= maxColumn; column++)
                {
                    for (var row = minRow; row <= maxRow; row++)
                    {
                        var id = new CellAddress(column, row).ToId();

                        if (sheet[id] is Subtotal)
                        {
                            continue;
                        }

                        values.Add(context.Workbook.GetCellValue(range.SheetName, id));
                    }
                }

                break;
            }

            case OpenRangeReference open:
            {
                var sheet = context.Workbook.Sheets[open.SheetName];

                foreach (var id in open.PopulatedIds(context))
                {
                    if (sheet[id] is not Subtotal)
                    {
                        values.Add(context.Workbook.GetCellValue(open.SheetName, id));
                    }
                }

                break;
            }

            case CellReference cell:
                if (context.Workbook.Sheets[cell.SheetName][cell.Id] is not Subtotal)
                {
                    values.Add(cell.Evaluate(context));
                }

                break;

            case UnionReference union:
                foreach (var area in union.Areas)
                {
                    GatherSkippingSubtotals(area, context, values);
                }

                break;

            default:
                var computed = argument.Evaluate(context);

                // A reference produced by a function (OFFSET, CHOOSE, …) still carries the actual
                // Reference node, so the nested-subtotal skip applies to it as well.
                if (computed.TryGetReference(out var reference))
                {
                    GatherSkippingSubtotals(reference, context, values);
                }
                else
                {
                    values.Add(computed);
                }

                break;
        }
    }

    private static ComputedValue Aggregate(int code, List<ComputedValue> values)
    {
        // COUNT (2) and COUNTA (3) never propagate cell errors, like their standalone functions.
        if (code == 2)
        {
            var count = 0;

            foreach (var value in values)
            {
                if (value.Kind == ComputedValueKind.Number)
                {
                    count++;
                }
            }

            return ComputedValue.Number(count);
        }

        if (code == 3)
        {
            var count = 0;

            foreach (var value in values)
            {
                if (value.Kind != ComputedValueKind.Blank)
                {
                    count++;
                }
            }

            return ComputedValue.Number(count);
        }

        var numbers = new List<double>();

        foreach (var value in values)
        {
            if (value.TryGetError(out var cellError))
            {
                return ComputedValue.Error(cellError);
            }

            // Referenced semantics: text, logicals and blanks are ignored.
            if (value.TryGetNumber(out var number))
            {
                numbers.Add(number);
            }
        }

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
