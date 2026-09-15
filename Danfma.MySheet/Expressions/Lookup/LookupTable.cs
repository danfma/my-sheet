using System.Diagnostics.CodeAnalysis;

namespace Danfma.MySheet.Expressions.Lookup;

internal static class LookupTable
{
    public static bool TryResolveTable(
        Expression[] arguments,
        EvaluationContext context,
        [NotNullWhen(true)] out Reference? reference,
        out ComputedValue answer
    )
    {
        if (!NamedReferences.TryResolveReference(arguments[1], context, out reference))
        {
            var tableValue = arguments[1].Evaluate(context);

            if (
                arguments[1] is not NameReference and not TableReference
                && !ArrayEvaluation.IsArrayEligible(arguments[1], context)
                && tableValue.TryGetError(out _)
            )
            {
                answer = ComputedValue.Error(Error.NA);
                return false;
            }

            answer = tableValue is { Kind: not ComputedValueKind.Error }
                ? LookupScalarTable(arguments, tableValue, context)
                : ReferencePosition.Unresolved(
                    arguments[1],
                    context,
                    ComputedValue.Error(Error.Ref)
                );
            return false;
        }

        if (reference is not CellReference cell)
        {
            answer = default;
            return true;
        }

        var value = cell.Evaluate(context);
        if (arguments[1] is NameReference && value.TryGetError(out _))
        {
            answer = value;
            return false;
        }

        answer = value.TryGetError(out _)
            ? DirectErrorCellTable(arguments, context)
            : LookupScalarTable(arguments, value, context);
        reference = null;
        return false;
    }

    public static ComputedValue LookupScalarTable(
        Expression[] arguments,
        ComputedValue tableValue,
        EvaluationContext context
    )
    {
        var lookup = arguments[0].Evaluate(context);
        if (ReferencePosition.IsLookupValueError(arguments[0], lookup, context, out var valueError))
        {
            return valueError;
        }

        if (arguments[2].Evaluate(context).CoerceToNumber(out var index) is { } indexError)
        {
            return ComputedValue.Error(indexError);
        }

        if (index < 1)
        {
            return ComputedValue.Error(Error.Value);
        }

        if (index > 1)
        {
            return ComputedValue.Error(Error.Ref);
        }

        var approximate = true;
        if (
            arguments.Length == 4
            && arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        return lookup.TryGetError(out _) || !ValueCoercion.AreEqual(tableValue, lookup)
            ? ComputedValue.Error(Error.NA)
            : tableValue;
    }

    public static ComputedValue DirectErrorCellTable(
        Expression[] arguments,
        EvaluationContext context
    )
    {
        var approximate = true;
        if (
            arguments.Length == 4
            && arguments[3].Evaluate(context).CoerceToBool(out approximate) is { } modeError
        )
        {
            return ComputedValue.Error(modeError);
        }

        return approximate ? arguments[1].Evaluate(context) : ComputedValue.Error(Error.NA);
    }
}
