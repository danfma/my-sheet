namespace Danfma.MySheet.Expressions.Logical;

/// <summary>
/// Shared argument reduction for the variadic logical functions AND / OR / XOR. Implements the Excel
/// rule for their operands: text and empty cells reached THROUGH a reference or array argument (a single
/// cell, a range, an open range, a union, or a cross-sheet reference) are IGNORED; only booleans and
/// numbers count (a number ≠ 0 is TRUE). A call whose arguments contribute no evaluable logical value at
/// all yields <see cref="Error.Value"/> (<c>#VALUE!</c>), matching the documented "no logical values ->
/// #VALUE!" remark on the AND/OR/XOR support pages.
/// </summary>
internal static class LogicalReduction
{
    /// <summary>
    /// Folds <paramref name="arguments"/> into the count of TRUE logical values (<paramref name="trueCount"/>)
    /// and the count of evaluable logical values (<paramref name="total"/>). Returns the <see cref="Error"/>
    /// to propagate (an error operand, or a non-numeric literal), or <c>null</c> on success. When
    /// <paramref name="total"/> is 0 the caller must return <c>#VALUE!</c>: nothing was evaluable.
    /// </summary>
    public static Error? Reduce(
        Expression[] arguments,
        EvaluationContext context,
        out int trueCount,
        out int total
    )
    {
        trueCount = 0;
        total = 0;

        foreach (var argument in arguments)
        {
            // A reference/array argument (cell, range, open range, union, cross-sheet): Excel ignores its
            // text and blank cells and evaluates only the logical/numeric entries.
            if (argument is Reference)
            {
                if (Accumulate(ArgumentFlattening.ExpandComputedValues(argument, context), ref trueCount, ref total)
                    is { } referenceError)
                {
                    return referenceError;
                }

                continue;
            }

            var computed = argument.Evaluate(context);

            // A function that yields a reference (OFFSET/INDEX/CHOOSE of a range, …) follows the same
            // ignore-text rule as a literal reference.
            if (computed.Kind == ComputedValueKind.Reference)
            {
                if (Accumulate(computed.EnumerateValues(context), ref trueCount, ref total) is { } valueError)
                {
                    return valueError;
                }

                continue;
            }

            // A direct scalar argument (a literal or the result of a comparison/arithmetic) is coerced
            // strictly: a non-numeric text LITERAL is #VALUE!, NOT ignored. This is a deliberately distinct
            // path from the reference/array case above — whether a literal text arg should be ignored is an
            // open oracle question (see plans/function-coverage-roadmap.md), so the literal keeps the 2.9.0
            // behavior until validated against real Excel.
            if (computed.CoerceToBool(out var value) is { } scalarError)
            {
                return scalarError;
            }

            total++;
            trueCount += value ? 1 : 0;
        }

        return null;
    }

    private static Error? Accumulate(
        IEnumerable<ComputedValue> values,
        ref int trueCount,
        ref int total
    )
    {
        foreach (var value in values)
        {
            switch (value.Kind)
            {
                case ComputedValueKind.Error:
                    value.TryGetError(out var error);
                    return error;

                case ComputedValueKind.Number:
                    value.TryGetNumber(out var number);
                    total++;
                    trueCount += number != 0 ? 1 : 0;
                    break;

                case ComputedValueKind.Boolean:
                    value.TryGetBoolean(out var boolean);
                    total++;
                    trueCount += boolean ? 1 : 0;
                    break;

                // Text and blank cells reached through a reference are ignored, per the Excel docs.
            }
        }

        return null;
    }
}
