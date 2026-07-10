using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

[MemoryPackable]
public sealed partial record Irr(Expression[] Arguments) : Function
{
    // IRR(values, [guess]) — the internal rate of return: the rate at which the cash flows (the
    // first at period 0) have a net present value of zero. Solved iteratively from `guess`
    // (default 0.1). #NUM! when the flows never change sign or the solver fails to converge.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A missing-sheet values range is a structural #REF! — an open-range ghost would otherwise be read as
        // empty and reported as #NUM! (no sign change).
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var flows = new List<double>();

        foreach (var value in ArgumentFlattening.ExpandComputedValues(Arguments[0], context))
        {
            if (value.TryGetError(out var error))
            {
                return ComputedValue.Error(error);
            }

            if (value.TryGetNumber(out var number))
            {
                flows.Add(number);
            }

            // Text, logicals and blanks are ignored, matching Excel.
        }

        var guess = 0.1;
        if (
            Arguments.Length > 1
            && Arguments[1].Evaluate(context).CoerceToNumber(out guess) is { } guessError
        )
        {
            return ComputedValue.Error(guessError);
        }

        if (!HasSignChange(flows))
        {
            return ComputedValue.Error(Error.Num);
        }

        // The IRR is the root of the period-0 NPV: Σ flow_t / (1+rate)^t.
        double Npv(double rate)
        {
            var sum = 0.0;
            var discount = 1.0;

            foreach (var flow in flows)
            {
                sum += flow / discount;
                discount *= 1 + rate;
            }

            return sum;
        }

        var result = TimeValueOfMoney.Solve(Npv, guess);
        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }

    private static bool HasSignChange(List<double> flows)
    {
        var hasPositive = false;
        var hasNegative = false;

        foreach (var flow in flows)
        {
            if (flow > 0)
            {
                hasPositive = true;
            }
            else if (flow < 0)
            {
                hasNegative = true;
            }
        }

        return hasPositive && hasNegative;
    }
}
