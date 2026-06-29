using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Ipmt(Expression[] Arguments) : Function
{
    // IPMT(rate, per, nper, pv, [fv], [type]) — the interest portion of the payment in period `per`.
    public override object? Compute(EvaluationContext context)
    {
        if (ValueCoercion.TryToNumber(Arguments[0].Compute(context), out var rate) is { } rateError)
        {
            return rateError;
        }

        if (ValueCoercion.TryToNumber(Arguments[1].Compute(context), out var per) is { } perError)
        {
            return perError;
        }

        if (ValueCoercion.TryToNumber(Arguments[2].Compute(context), out var nper) is { } nperError)
        {
            return nperError;
        }

        if (ValueCoercion.TryToNumber(Arguments[3].Compute(context), out var pv) is { } pvError)
        {
            return pvError;
        }

        var fv = 0.0;
        if (
            Arguments.Length > 4
            && ValueCoercion.TryToNumber(Arguments[4].Compute(context), out fv) is { } fvError
        )
        {
            return fvError;
        }

        var type = 0.0;
        if (
            Arguments.Length > 5
            && ValueCoercion.TryToNumber(Arguments[5].Compute(context), out type) is { } typeError
        )
        {
            return typeError;
        }

        if (per < 1 || per > nper)
        {
            return ErrorValue.Number;
        }

        var result = TimeValueOfMoney.IPmt(rate, per, nper, pv, fv, type != 0 ? 1 : 0);
        return double.IsFinite(result) ? result : ErrorValue.Number;
    }
}
