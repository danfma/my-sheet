namespace MySheet.Expressions;

/// <summary>
/// Flattens function arguments into a sequence of computed values, expanding range arguments
/// cell-by-cell. Used by variadic functions (CONCAT, TEXTJOIN, COUNTA, …).
/// </summary>
internal static class ArgumentFlattening
{
    public static IEnumerable<object?> Flatten(Expression[] arguments, EvaluationContext context)
    {
        foreach (var argument in arguments)
        {
            if (argument is RangeReference range)
            {
                foreach (var value in range.ExpandValues(context))
                {
                    yield return value;
                }
            }
            else if (argument is UnionReference union)
            {
                foreach (var value in union.ExpandValues(context))
                {
                    yield return value;
                }
            }
            else
            {
                var value = argument.Compute(context);

                if (value is RangeReference resultRange)
                {
                    foreach (var cellValue in resultRange.ExpandValues(context))
                    {
                        yield return cellValue;
                    }
                }
                else
                {
                    yield return value;
                }
            }
        }
    }

    /// <summary>
    /// Expands a single argument into an ordered list of computed values (a range yields one entry per
    /// cell). Used by the conditional aggregations to align parallel ranges by position.
    /// </summary>
    public static List<object?> Expand(Expression argument, EvaluationContext context)
    {
        var values = new List<object?>();

        if (argument is RangeReference range)
        {
            foreach (var value in range.ExpandValues(context))
            {
                values.Add(value);
            }
        }
        else if (argument is UnionReference union)
        {
            foreach (var value in union.ExpandValues(context))
            {
                values.Add(value);
            }
        }
        else
        {
            var value = argument.Compute(context);

            if (value is RangeReference resultRange)
            {
                foreach (var cellValue in resultRange.ExpandValues(context))
                {
                    values.Add(cellValue);
                }
            }
            else
            {
                values.Add(value);
            }
        }

        return values;
    }
}
