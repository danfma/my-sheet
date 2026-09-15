namespace Danfma.MySheet.Expressions.Lookup;

/// <summary>
/// Builds LOOKUP's vector slots exactly once after selecting scalar <c>IF</c>/<c>CHOOSE</c> paths.
/// Measured Aspose.Cells 26.7.0 rule: lookup-slot scalar errors propagate; element errors remain in a
/// materialized vector and are skipped by matching; a result-slot singleton error against multiple keys is
/// <c>#N/A</c>; structural missing sheets on the selected path are always <c>#REF!</c>.
/// </summary>
internal static class LookupVector
{
    public static bool TryMaterializeLookup(
        Expression expression,
        EvaluationContext context,
        out IReadOnlyList<ComputedValue> values,
        out ComputedValue error
    )
    {
        if (TrySelectedPathError(expression, context, propagateOwnError: true, out error))
        {
            values = [];
            return false;
        }

        values = ArgumentFlattening.MaterializeVector(expression, context);
        return true;
    }

    public static bool TryMaterializeResult(
        Expression expression,
        EvaluationContext context,
        int lookupCount,
        out IReadOnlyList<ComputedValue> values,
        out ComputedValue error
    )
    {
        if (TrySelectedPathError(expression, context, propagateOwnError: false, out error))
        {
            values = [];
            return false;
        }

        values = ArgumentFlattening.MaterializeVector(expression, context);
        if (lookupCount > 1 && values is [var only] && only.TryGetError(out _))
        {
            error = ComputedValue.Error(Error.NA);
            return false;
        }

        return true;
    }

    private static bool TrySelectedPathError(
        Expression expression,
        EvaluationContext context,
        bool propagateOwnError,
        out ComputedValue error
    )
    {
        error = ComputedValue.Blank;
        if (ReferenceGuard.MissingSheet(expression, context) is { } missing)
        {
            error = ComputedValue.Error(missing);
            return true;
        }

        switch (expression)
        {
            case Logical.If ifNode when ifNode.Arguments.Length is 2 or 3:
                var conditionError = context
                    .EvaluateConditionOnce(ifNode.Arguments[0])
                    .CoerceToBoolAllowingTextWords(out var condition);
                if (conditionError is not null)
                {
                    if (propagateOwnError)
                    {
                        error = ComputedValue.Error(conditionError.Value);
                    }
                    return propagateOwnError;
                }

                var selected =
                    condition ? ifNode.Arguments[1]
                    : ifNode.Arguments.Length == 3 ? ifNode.Arguments[2]
                    : null;
                return selected is not null
                    && TrySelectedPathError(selected, context, propagateOwnError, out error);

            case Choose choose:
                if (choose.TryChoose(context, out var chosen) is { } chooseError)
                {
                    if (propagateOwnError)
                    {
                        error = chooseError;
                    }
                    return propagateOwnError;
                }

                return TrySelectedPathError(chosen, context, propagateOwnError, out error);

            case BinaryOperation binary
                when !ArrayEvaluation.IsArrayEligible(binary.Left, context)
                    || !ArrayEvaluation.IsArrayEligible(binary.Right, context):
                if (
                    TryStructuralError(binary.Left, context, out error)
                    || TryStructuralError(binary.Right, context, out error)
                )
                {
                    return true;
                }
                break;
        }

        if (
            propagateOwnError
            && !IsComputedVector(expression, context)
            && ReferencePosition.TryUnresolvedError(expression, context, out error)
        )
        {
            return true;
        }

        return false;
    }

    private static bool TryStructuralError(
        Expression expression,
        EvaluationContext context,
        out ComputedValue error
    ) => TrySelectedPathError(expression, context, propagateOwnError: false, out error);

    private static bool IsComputedVector(Expression expression, EvaluationContext context) =>
        expression is not TableReference
        && (ArrayEvaluation.IsArrayEligible(expression, context) || ContainsReference(expression));

    private static bool ContainsReference(Expression expression) =>
        expression switch
        {
            Reference => true,
            BinaryOperation binary => ContainsReference(binary.Left)
                || ContainsReference(binary.Right),
            UnaryOperation unary => ContainsReference(unary.Operand),
            _ => false,
        };
}
