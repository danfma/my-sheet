using System.Globalization;

namespace MySheet.Expressions;

/// <summary>
/// Shared numeric gathering for aggregate functions (SUM, AVERAGE, MIN, MAX, COUNT).
/// Mirrors Excel's rule: numeric text and logicals passed <em>directly</em> as arguments are counted,
/// but text/logicals/blanks pulled from <em>referenced</em> cells are ignored. The first error
/// encountered is reported; the caller decides whether to propagate it (SUM…) or ignore it (COUNT).
/// </summary>
internal static class NumericAggregation
{
    public static (List<double> Numbers, ErrorValue? Error) Gather(Expression[] arguments, Workbook workbook)
    {
        var numbers = new List<double>();
        ErrorValue? error = null;

        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var cell in range.Expand(workbook))
                    {
                        AddReferenced(cell.Compute(workbook), numbers, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferenced(cell.Compute(workbook), numbers, ref error);
                    break;

                default:
                    AddDirect(argument.Compute(workbook), numbers, ref error);
                    break;
            }
        }

        return (numbers, error);
    }

    private static void AddReferenced(object? value, List<double> numbers, ref ErrorValue? error)
    {
        switch (value)
        {
            case ErrorValue referencedError:
                error ??= referencedError;
                break;

            case double number:
                numbers.Add(number);
                break;

            // Referenced text, logicals and blanks are ignored, matching Excel.
        }
    }

    private static void AddDirect(object? value, List<double> numbers, ref ErrorValue? error)
    {
        switch (value)
        {
            case ErrorValue directError:
                error ??= directError;
                break;

            case double number:
                numbers.Add(number);
                break;

            case bool boolean:
                numbers.Add(boolean ? 1 : 0);
                break;

            case null:
                // Blank ignored.
                break;

            case string text
                when double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed):
                numbers.Add(parsed);
                break;

            case string:
                error ??= ErrorValue.NotValue;
                break;
        }
    }
}
