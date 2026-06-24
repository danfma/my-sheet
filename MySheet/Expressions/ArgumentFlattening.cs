namespace MySheet.Expressions;

/// <summary>
/// Flattens function arguments into a sequence of computed values, expanding range arguments
/// cell-by-cell. Used by variadic functions (CONCAT, TEXTJOIN, COUNTA, …).
/// </summary>
internal static class ArgumentFlattening
{
    public static IEnumerable<object?> Flatten(Expression[] arguments, Workbook workbook)
    {
        foreach (var argument in arguments)
        {
            if (argument is RangeReference range)
            {
                foreach (var cell in range.Expand(workbook))
                {
                    yield return cell.Compute(workbook);
                }
            }
            else
            {
                yield return argument.Compute(workbook);
            }
        }
    }

    /// <summary>
    /// Expands a single argument into an ordered list of computed values (a range yields one entry per
    /// cell). Used by the conditional aggregations to align parallel ranges by position.
    /// </summary>
    public static List<object?> Expand(Expression argument, Workbook workbook)
    {
        var values = new List<object?>();

        if (argument is RangeReference range)
        {
            foreach (var cell in range.Expand(workbook))
            {
                values.Add(cell.Compute(workbook));
            }
        }
        else
        {
            values.Add(argument.Compute(workbook));
        }

        return values;
    }
}
