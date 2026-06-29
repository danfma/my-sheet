using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Nper(Expression[] Arguments) : Function
{
    // NPER(rate, pmt, pv, [fv], [type]) — the number of periods of an annuity.
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var rate) is { } rateError)
        {
            return rateError;
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

        var result = TimeValueOfMoney.Nper(rate, pmt, pv, fv, type != 0 ? 1 : 0);
        return double.IsFinite(result) ? result : ErrorValue.Number;
    }
}
