using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Rate(Expression[] Arguments) : Function
{
    // RATE(nper, pmt, pv, [fv], [type], [guess]) — the periodic interest rate of an annuity. There
    // is no closed form, so it is solved iteratively from `guess` (default 0.1); #NUM! if it fails.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var nper) is { } nperError)
        {
            return ComputedValue.Error(nperError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var pmt) is { } pmtError)
        {
            return ComputedValue.Error(pmtError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var pv) is { } pvError)
        {
            return ComputedValue.Error(pvError);
        }

        var fv = 0.0;
        if (
            Arguments.Length > 3
            && Arguments[3].Evaluate(context).CoerceToNumber(out fv) is { } fvError
        )
        {
            return ComputedValue.Error(fvError);
        }

        var type = 0.0;
        if (
            Arguments.Length > 4
            && Arguments[4].Evaluate(context).CoerceToNumber(out type) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        var guess = 0.1;
        if (
            Arguments.Length > 5
            && Arguments[5].Evaluate(context).CoerceToNumber(out guess) is { } guessError
        )
        {
            return ComputedValue.Error(guessError);
        }

        var normalizedType = type != 0 ? 1 : 0;

        // RATE solves Fv(rate, …) = fv for rate; the root of this residual is the answer.
        double Residual(double rate) =>
            TimeValueOfMoney.Fv(rate, nper, pmt, pv, normalizedType) - fv;

        var result = TimeValueOfMoney.Solve(Residual, guess);
        return double.IsFinite(result) ? ComputedValue.Number(result) : ComputedValue.Error(Error.Num);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
