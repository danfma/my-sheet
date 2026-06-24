using System.Globalization;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Receives each numeric value gathered by <see cref="NumericAggregation"/>. Implemented by mutable
/// structs so the JIT specializes the fold per function and avoids heap allocation in the hot path.
/// </summary>
internal interface INumericFold
{
    void Accept(double value);
}

/// <summary>
/// Shared single-pass numeric gathering for aggregate functions (SUM, AVERAGE, MIN, MAX, COUNT).
/// Mirrors Excel's rule: numeric text and logicals passed <em>directly</em> as arguments are counted,
/// but text/logicals/blanks pulled from <em>referenced</em> cells are ignored. The first error
/// encountered is returned; the caller decides whether to propagate it (SUM…) or ignore it (COUNT).
/// </summary>
internal static class NumericAggregation
{
    public static ErrorValue? Fold<TFold>(
        Expression[] arguments,
        EvaluationContext context,
        ref TFold fold
    )
        where TFold : struct, INumericFold
    {
        ErrorValue? error = null;

        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferenced(cell.Compute(context), ref fold, ref error);
                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                default:
                    var argumentValue = argument.Compute(context);

                    // A function (e.g. OFFSET) may yield a range value; expand it as referenced cells.
                    if (argumentValue is RangeReference resultRange)
                    {
                        foreach (var cellValue in resultRange.ExpandValues(context))
                        {
                            AddReferenced(cellValue, ref fold, ref error);
                        }
                    }
                    else
                    {
                        AddDirect(argumentValue, ref fold, ref error);
                    }

                    break;
            }
        }

        return error;
    }

    private static void AddReferenced<TFold>(object? value, ref TFold fold, ref ErrorValue? error)
        where TFold : struct, INumericFold
    {
        switch (value)
        {
            case ErrorValue referencedError:
                error ??= referencedError;
                break;

            case double number:
                fold.Accept(number);
                break;

            // Referenced text, logicals and blanks are ignored, matching Excel.
        }
    }

    private static void AddDirect<TFold>(object? value, ref TFold fold, ref ErrorValue? error)
        where TFold : struct, INumericFold
    {
        switch (value)
        {
            case ErrorValue directError:
                error ??= directError;
                break;

            case double number:
                fold.Accept(number);
                break;

            case bool boolean:
                fold.Accept(boolean ? 1 : 0);
                break;

            case null:
                // Blank ignored.
                break;

            case string text
                when double.TryParse(
                    text,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out var parsed
                ):
                fold.Accept(parsed);
                break;

            case string:
                error ??= ErrorValue.NotValue;
                break;
        }
    }
}
