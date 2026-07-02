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
/// Fold that materializes the gathered values, for aggregations that need more than a single pass
/// (GCD, LCM, MULTINOMIAL).
/// </summary>
internal struct NumericListFold : INumericFold
{
    public List<double> Values;

    public void Accept(double value) => Values.Add(value);
}

/// <summary>
/// Shared single-pass numeric gathering for aggregate functions (SUM, AVERAGE, MIN, MAX, COUNT), reading
/// each cell as a <see cref="ComputedValue"/> straight from the cache (no boxing). Mirrors Excel's rule:
/// numeric text and logicals passed <em>directly</em> as arguments are counted, but text/logicals/blanks
/// pulled from <em>referenced</em> cells are ignored. The first error encountered is returned; the caller
/// decides whether to propagate it (SUM…) or ignore it (COUNT).
/// </summary>
internal static class NumericAggregation
{
    public static Error? Fold<TFold>(
        Expression[] arguments,
        EvaluationContext context,
        ref TFold fold
    )
        where TFold : struct, INumericFold
    {
        Error? error = null;

        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandComputedValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                case CellReference cell:
                    AddReferenced(cell.Evaluate(context), ref fold, ref error);
                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandComputedValues(context))
                    {
                        AddReferenced(value, ref fold, ref error);
                    }

                    break;

                default:
                    var argumentValue = argument.Evaluate(context);

                    // A function (e.g. OFFSET) may yield a range value; expand it as referenced cells.
                    if (argumentValue.Kind == ComputedValueKind.Reference)
                    {
                        foreach (var cellValue in argumentValue.EnumerateValues(context))
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

    private static void AddReferenced<TFold>(in ComputedValue value, ref TFold fold, ref Error? error)
        where TFold : struct, INumericFold
    {
        if (value.TryGetError(out var referencedError))
        {
            error ??= referencedError;
        }
        else if (value.TryGetNumber(out var number))
        {
            fold.Accept(number);
        }

        // Referenced text, logicals and blanks are ignored, matching Excel.
    }

    private static void AddDirect<TFold>(in ComputedValue value, ref TFold fold, ref Error? error)
        where TFold : struct, INumericFold
    {
        switch (value.Kind)
        {
            case ComputedValueKind.Error:
                value.TryGetError(out var directError);
                error ??= directError;
                break;

            case ComputedValueKind.Number:
                value.TryGetNumber(out var number);
                fold.Accept(number);
                break;

            case ComputedValueKind.Boolean:
                value.TryGetBoolean(out var boolean);
                fold.Accept(boolean ? 1 : 0);
                break;

            case ComputedValueKind.Blank:
                // Blank ignored.
                break;

            case ComputedValueKind.Text:
                value.TryGetText(out var text);
                if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                {
                    fold.Accept(parsed);
                }
                else
                {
                    error ??= Error.Value;
                }

                break;
        }
    }
}
