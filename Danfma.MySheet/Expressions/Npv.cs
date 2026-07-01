using System.Globalization;
using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Npv(Expression[] Arguments) : Function
{
    // NPV(rate, value1, [value2], …) — net present value of a series of cash flows, the first
    // discounted by one period. Like Excel, numeric text and logicals passed *directly* count, but
    // text/logicals/blanks pulled from *referenced* cells are ignored.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
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
        ErrorValue? error = null;

        for (var i = 1; i < Arguments.Length && error is null; i++)
        {
            switch (Arguments[i])
            {
                case RangeReference range:
                    foreach (var value in range.ExpandValues(context))
                    {
                        AddReferenced(value, denominator, ref discount, ref sum, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferenced(cell.Compute(context), denominator, ref discount, ref sum, ref error);
                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandValues(context))
                    {
                        AddReferenced(value, denominator, ref discount, ref sum, ref error);
                    }

                    break;

                default:
                    var argumentValue = Arguments[i].Compute(context);

                    if (argumentValue is RangeReference resultRange)
                    {
                        foreach (var cellValue in resultRange.ExpandValues(context))
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

        return error is { } propagated ? ComputedValue.From(propagated) : ComputedValue.Number(sum);
    }

    private static void AddReferenced(
        object? value,
        double denominator,
        ref double discount,
        ref double sum,
        ref ErrorValue? error
    )
    {
        switch (value)
        {
            case ErrorValue referencedError:
                error ??= referencedError;
                break;

            case double number:
                discount *= denominator;
                sum += number / discount;
                break;

            // Referenced text, logicals and blanks are ignored, matching Excel.
        }
    }

    private static void AddDirect(
        object? value,
        double denominator,
        ref double discount,
        ref double sum,
        ref ErrorValue? error
    )
    {
        switch (value)
        {
            case ErrorValue directError:
                error ??= directError;
                break;

            case double number:
                discount *= denominator;
                sum += number / discount;
                break;

            case bool boolean:
                discount *= denominator;
                sum += (boolean ? 1 : 0) / discount;
                break;

            case null:
                break;

            case string text
                when double.TryParse(
                    text,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out var parsed
                ):
                discount *= denominator;
                sum += parsed / discount;
                break;

            case string:
                error ??= ErrorValue.NotValue;
                break;
        }
    }
}
