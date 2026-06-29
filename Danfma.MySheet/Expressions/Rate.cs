using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Rate(Expression[] Arguments) : Function
{
    // RATE(nper, pmt, pv, [fv], [type], [guess]) — the periodic interest rate of an annuity. There
    // is no closed form, so it is solved iteratively from `guess` (default 0.1); #NUM! if it fails.
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var nper) is { } nperError)
        {
            return nperError;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(context), out var pmt) is { } pmtError)
        {
            return pmtError;
        }

        if (ValueCoercion.TryToNumber(Arguments[2].Compute(context), out var pv) is { } pvError)
        {
            return pvError;
        }

        var fv = 0.0;
        if (
            Arguments.Length > 3
            && ValueCoercion.TryToNumber(Arguments[3].Compute(context), out fv) is { } fvError
        )
        {
            return fvError;
        }

        var type = 0.0;
        if (
            Arguments.Length > 4
            && ValueCoercion.TryToNumber(Arguments[4].Compute(context), out type) is { } typeError
        )
        {
            return typeError;
        }

        var guess = 0.1;
        if (
            Arguments.Length > 5
            && ValueCoercion.TryToNumber(Arguments[5].Compute(context), out guess) is { } guessError
        )
        {
            return guessError;
        }

        var normalizedType = type != 0 ? 1 : 0;

        // RATE solves Fv(rate, …) = fv for rate; the root of this residual is the answer.
        double Residual(double rate) =>
            TimeValueOfMoney.Fv(rate, nper, pmt, pv, normalizedType) - fv;

        var result = TimeValueOfMoney.Solve(Residual, guess);
        return double.IsFinite(result) ? result : ErrorValue.Number;
    }
}
