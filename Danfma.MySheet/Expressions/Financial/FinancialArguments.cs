namespace Danfma.MySheet.Expressions.Financial;

/// <summary>
/// Shared argument coercion and domain validation for the wave-6 financial functions. Date arguments arrive
/// as numeric serials (Excel truncates the time of day), frequency must be 1/2/4 and day-count basis 0..4;
/// anything out of range maps to <see cref="Error.Num"/>, matching Excel.
/// </summary>
internal static class FinancialArguments
{
    public static Error? Number(Expression[] arguments, int index, EvaluationContext context, out double value) =>
        arguments[index].Evaluate(context).CoerceToNumber(out value);

    /// <summary>Coerces the argument to a numeric serial and converts it to a whole-day date.</summary>
    public static Error? Date(Expression[] arguments, int index, EvaluationContext context, out DateTime date)
    {
        date = default;

        if (arguments[index].Evaluate(context).CoerceToNumber(out var serial) is { } error)
        {
            return error;
        }

        return DateSerial.ToDateTime(Math.Floor(serial), out date);
    }

    /// <summary>Validates the coupon frequency argument: only 1 (annual), 2 (semi-annual), 4 (quarterly).</summary>
    public static Error? Frequency(Expression[] arguments, int index, EvaluationContext context, out int frequency)
    {
        frequency = 0;

        if (arguments[index].Evaluate(context).CoerceToNumber(out var raw) is { } error)
        {
            return error;
        }

        var truncated = (int)Math.Truncate(raw);
        if (truncated is not (1 or 2 or 4))
        {
            return Error.Num;
        }

        frequency = truncated;
        return null;
    }

    /// <summary>Validates the day-count basis argument (0..4). Absent argument defaults to 0.</summary>
    public static Error? Basis(
        Expression[] arguments,
        int index,
        EvaluationContext context,
        out int basis
    )
    {
        basis = 0;

        if (arguments.Length <= index)
        {
            return null;
        }

        if (arguments[index].Evaluate(context).CoerceToNumber(out var raw) is { } error)
        {
            return error;
        }

        var truncated = (int)Math.Truncate(raw);
        if (truncated is < 0 or > 4)
        {
            return Error.Num;
        }

        basis = truncated;
        return null;
    }

    /// <summary>Wraps a finite double result, mapping NaN/Infinity to <see cref="Error.Num"/>.</summary>
    public static ComputedValue Result(double value) =>
        double.IsFinite(value) ? ComputedValue.Number(value) : ComputedValue.Error(Error.Num);

    /// <summary>
    /// Reads the (settlement, maturity, frequency, [basis]) prefix shared by every coupon function and
    /// validates the common bond domain: settlement &lt; maturity, frequency ∈ {1,2,4}, basis ∈ {0..4}.
    /// </summary>
    public static Error? CouponInputs(
        Expression[] arguments,
        EvaluationContext context,
        out DateTime settlement,
        out DateTime maturity,
        out int frequency,
        out int basis
    )
    {
        maturity = default;
        frequency = 0;
        basis = 0;

        if (Date(arguments, 0, context, out settlement) is { } settlementError)
        {
            return settlementError;
        }

        if (Date(arguments, 1, context, out maturity) is { } maturityError)
        {
            return maturityError;
        }

        if (Frequency(arguments, 2, context, out frequency) is { } frequencyError)
        {
            return frequencyError;
        }

        if (Basis(arguments, 3, context, out basis) is { } basisError)
        {
            return basisError;
        }

        return settlement >= maturity ? Error.Num : null;
    }

    /// <summary>
    /// Flattens an argument (expanding ranges) into a list of numbers, ignoring text/logical/blank cells the
    /// way IRR and NPV do. Returns the first error encountered, or <c>null</c> on success.
    /// </summary>
    public static Error? NumberList(Expression argument, EvaluationContext context, List<double> values)
    {
        foreach (var value in ArgumentFlattening.ExpandComputedValues(argument, context))
        {
            if (value.TryGetError(out var error))
            {
                return error;
            }

            if (value.TryGetNumber(out var number))
            {
                values.Add(number);
            }
        }

        return null;
    }
}
