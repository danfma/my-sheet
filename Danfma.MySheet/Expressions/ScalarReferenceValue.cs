namespace Danfma.MySheet.Expressions;

internal static class ScalarReferenceValue
{
    public static ComputedValue Evaluate(Expression expression, EvaluationContext context)
    {
        if (
            NamedReferences.TryResolveReference(
                expression,
                context,
                out var reference,
                out var unresolvedValue,
                boundOpenRanges: false
            )
        )
        {
            if (reference is CellReference cell)
            {
                return cell.Evaluate(context);
            }

            return ImplicitIntersection.Apply(reference, context);
        }

        return unresolvedValue ?? expression.Evaluate(context);
    }
}
