using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

[MemoryPackable]
public sealed partial record Npv(Expression[] Arguments) : Function
{
    // NPV(rate, value1, [value2], …) — net present value of a series of cash flows, the first
    // discounted by one period. Like Excel, numeric text and logicals passed *directly* count, but
    // text/logicals/blanks pulled from *referenced* cells are ignored.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // Any reference argument (the rate cell or a cash-flow range) to a missing sheet is a structural
        // #REF! — an open-range ghost would otherwise be swallowed as empty, silently yielding 0.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        if (Arguments[0].Evaluate(context).CoerceToNumber(out var rate) is { } rateError)
        {
            return ComputedValue.Error(rateError);
        }

        var denominator = 1 + rate;
        if (denominator == 0)
        {
            return ComputedValue.Error(Error.DivZero);
        }

        var sum = 0.0;
        var discount = 1.0; // (1+rate)^period; multiplied by the denominator before each cash flow.
        Error? error = null;

        for (var i = 1; i < Arguments.Length && error is null; i++)
        {
            switch (Arguments[i])
            {
                case RangeReference range:
                    foreach (var value in range.ExpandComputedValues(context))
                    {
                        AddReferenced(value, denominator, ref discount, ref sum, ref error);
                    }

                    break;

                case OpenRangeReference open:
                    foreach (var value in open.ExpandComputedValues(context))
                    {
                        AddReferenced(value, denominator, ref discount, ref sum, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferenced(cell.Evaluate(context), denominator, ref discount, ref sum, ref error);
                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandComputedValues(context))
                    {
                        AddReferenced(value, denominator, ref discount, ref sum, ref error);
                    }

                    break;

                default:
                    var argumentValue = Arguments[i].Evaluate(context);

                    if (argumentValue.Kind == ComputedValueKind.Reference)
                    {
                        foreach (var cellValue in argumentValue.EnumerateValues(context))
                        {
                            AddReferenced(cellValue, denominator, ref discount, ref sum, ref error);
                        }
                    }
                    else
                    {
                        AddDirect(argumentValue, denominator, ref discount, ref sum, ref error);
                    }

                    break;
            }
        }

        return error is { } propagated ? ComputedValue.Error(propagated) : ComputedValue.Number(sum);
    }

    private static void AddReferenced(
        in ComputedValue value,
        double denominator,
        ref double discount,
        ref double sum,
        ref Error? error
    )
    {
        if (value.TryGetError(out var referencedError))
        {
            error ??= referencedError;
        }
        else if (value.TryGetNumber(out var number))
        {
            discount *= denominator;
            sum += number / discount;
        }

        // Referenced text, logicals and blanks are ignored, matching Excel.
    }

    private static void AddDirect(
        in ComputedValue value,
        double denominator,
        ref double discount,
        ref double sum,
        ref Error? error
    )
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Error:
                value.TryGetError(out var directError);
                error ??= directError;
                break;

            case ComputedValueKind.Number:
                value.TryGetNumber(out var number);
                discount *= denominator;
                sum += number / discount;
                break;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out var boolean);
                discount *= denominator;
                sum += (boolean ? 1 : 0) / discount;
                break;

            case ComputedValueKind.Blank:
                break;

            case ComputedValueKind.Text:
                value.TryGetText(out var text);
                if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                {
                    discount *= denominator;
                    sum += parsed / discount;
                }
                else
                {
                    error ??= Error.Value;
                }

                break;
        }
    }
}
