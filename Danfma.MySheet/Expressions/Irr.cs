using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Irr(Expression[] Arguments) : Function
{
    // IRR(values, [guess]) — the internal rate of return: the rate at which the cash flows (the
    // first at period 0) have a net present value of zero. Solved iteratively from `guess`
    // (default 0.1). #NUM! when the flows never change sign or the solver fails to converge.
    public override object? Compute(EvaluationContext context)
    {
        var flows = new List<double>();

        foreach (var value in ArgumentFlattening.Expand(Arguments[0], context))
        {
            switch (value)
            {
                case ErrorValue error:
                    return error;

                case double number:
                    flows.Add(number);
                    break;

                // Text, logicals and blanks are ignored, matching Excel.
            }
        }

        var guess = 0.1;
        if (
            Arguments.Length > 1
            && ValueCoercion.TryToNumber(Arguments[1].Compute(context), out guess) is { } guessError
        )
        {
            return guessError;
        }

        if (!HasSignChange(flows))
        {
            return ErrorValue.Number;
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
        return double.IsFinite(result) ? result : ErrorValue.Number;
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
