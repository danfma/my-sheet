using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Fv(Expression[] Arguments) : Function
{
    // FV(rate, nper, pmt, [pv], [type]) — the future value of an annuity.
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var rate) is { } rateError)
        {
            return rateError;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(context), out var nper) is { } nperError)
        {
            return nperError;
        }

        if (ValueCoercion.TryToNumber(Arguments[2].Compute(context), out var pmt) is { } pmtError)
        {
            return pmtError;
        }

        var pv = 0.0;
        if (
            Arguments.Length > 3
            && ValueCoercion.TryToNumber(Arguments[3].Compute(context), out pv) is { } pvError
        )
        {
            return pvError;
        }

        var type = 0.0;
        if (
            Arguments.Length > 4
            && ValueCoercion.TryToNumber(Arguments[4].Compute(context), out type) is { } typeError
        )
        {
            return typeError;
        }

        var result = TimeValueOfMoney.Fv(rate, nper, pmt, pv, type != 0 ? 1 : 0);
        return double.IsFinite(result) ? result : ErrorValue.Number;
    }
}
