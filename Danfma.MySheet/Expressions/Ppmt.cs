using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Ppmt(Expression[] Arguments) : Function
{
    // PPMT(rate, per, nper, pv, [fv], [type]) — the principal portion of the payment in period
    // `per`, i.e. the total payment minus its interest portion.
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

        var normalizedType = type != 0 ? 1 : 0;
        var payment = TimeValueOfMoney.Pmt(rate, nper, pv, fv, normalizedType);
        var interest = TimeValueOfMoney.IPmt(rate, per, nper, pv, fv, normalizedType);
        var result = payment - interest;
        return double.IsFinite(result) ? result : ErrorValue.Number;
    }
}
