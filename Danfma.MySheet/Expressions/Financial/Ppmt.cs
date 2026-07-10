using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

[MemoryPackable]
public sealed partial record Ppmt(Expression[] Arguments) : Function
{
    // PPMT(rate, per, nper, pv, [fv], [type]) — the principal portion of the payment in period
    // `per`, i.e. the total payment minus its interest portion.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var rate) is { } rateError)
        {
            return ComputedValue.Error(rateError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var per) is { } perError)
        {
            return ComputedValue.Error(perError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var nper) is { } nperError)
        {
            return ComputedValue.Error(nperError);
        }

        if (Arguments[3].Evaluate(context).CoerceToNumber(out var pv) is { } pvError)
        {
            return ComputedValue.Error(pvError);
        }

        var fv = 0.0;
        if (
            Arguments.Length > 4
            && Arguments[4].Evaluate(context).CoerceToNumber(out fv) is { } fvError
        )
        {
            return ComputedValue.Error(fvError);
        }

        var type = 0.0;
        if (
            Arguments.Length > 5
            && Arguments[5].Evaluate(context).CoerceToNumber(out type) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        if (per < 1 || per > nper)
        {
            return ComputedValue.Error(Error.Num);
        }

        var normalizedType = type != 0 ? 1 : 0;
        var payment = TimeValueOfMoney.Pmt(rate, nper, pv, fv, normalizedType);
        var interest = TimeValueOfMoney.IPmt(rate, per, nper, pv, fv, normalizedType);
        var result = payment - interest;
        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}
