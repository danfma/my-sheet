namespace Danfma.MySheet.Expressions;

/// <summary>
/// Flattens function arguments into a sequence of computed values, expanding range arguments
/// cell-by-cell. Used by variadic functions (CONCAT, TEXTJOIN, COUNTA, …). The <see cref="ComputedValue"/>
/// overloads read straight from the cache (no boxing); the <c>object?</c> ones are boxed views for interop.
/// </summary>
internal static class ArgumentFlattening
{
    public static IEnumerable<ComputedValue> FlattenComputedValues(
        Expression[] arguments,
        EvaluationContext context
    )
    {
        foreach (var argument in arguments)
        {
            switch (argument)
            {
                case RangeReference range:
                    foreach (var value in range.ExpandComputedValues(context))
                    {
                        yield return value;
                    }

                    break;

                case UnionReference union:
                    foreach (var value in union.ExpandComputedValues(context))
                    {
                        yield return value;
                    }

                    break;

                default:
                    var computed = argument.Evaluate(context);

                    if (computed.Kind == ComputedValueKind.Reference)
                    {
                        foreach (var cellValue in computed.EnumerateValues(context))
                        {
                            yield return cellValue;
                        }
                    }
                    else
                    {
                        yield return computed;
                    }

                    break;
            }
        }
    }

    /// <summary>
    /// Expands a single argument into an ordered list of computed values (a range yields one entry per
    /// cell). Used by the conditional aggregations to align parallel ranges by position.
    /// </summary>
    public static List<ComputedValue> ExpandComputedValues(Expression argument, EvaluationContext context)
    {
        var values = new List<ComputedValue>();

        switch (argument)
        {
            case RangeReference range:
                foreach (var value in range.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                break;

            case UnionReference union:
                foreach (var value in union.ExpandComputedValues(context))
                {
                    values.Add(value);
                }

                break;

            default:
                var computed = argument.Evaluate(context);

                if (computed.Kind == ComputedValueKind.Reference)
                {
                    foreach (var cellValue in computed.EnumerateValues(context))
                    {
                        values.Add(cellValue);
                    }
                }
                else
                {
                    values.Add(computed);
                }

                break;
        }

        return values;
    }

    /// <summary>Boxed (<c>object?</c>) view of <see cref="FlattenComputedValues"/>, for interop.</summary>
    public static IEnumerable<object?> Flatten(Expression[] arguments, EvaluationContext context)
    {
        foreach (var value in FlattenComputedValues(arguments, context))
        {
            yield return value.AsObject();
        }
    }

    /// <summary>Boxed (<c>object?</c>) view of <see cref="ExpandComputedValues"/>, for interop.</summary>
    public static List<object?> Expand(Expression argument, EvaluationContext context)
    {
        var values = new List<object?>();

        foreach (var value in ExpandComputedValues(argument, context))
        {
            values.Add(value.AsObject());
        }

        return values;
    }
}
