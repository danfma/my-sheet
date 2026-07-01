using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Fv(Expression[] Arguments) : Function
{
    // FV(rate, nper, pmt, [pv], [type]) — the future value of an annuity.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var rate) is { } rateError)
        {
            return ComputedValue.Error(rateError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var nper) is { } nperError)
        {
            return ComputedValue.Error(nperError);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var pmt) is { } pmtError)
        {
            return ComputedValue.Error(pmtError);
        }

        var pv = 0.0;
        if (
            Arguments.Length > 3
            && Arguments[3].Evaluate(context).CoerceToNumber(out pv) is { } pvError
        )
        {
            return ComputedValue.Error(pvError);
        }

        var type = 0.0;
        if (
            Arguments.Length > 4
            && Arguments[4].Evaluate(context).CoerceToNumber(out type) is { } typeError
        )
        {
            return ComputedValue.Error(typeError);
        }

        var result = TimeValueOfMoney.Fv(rate, nper, pmt, pv, type != 0 ? 1 : 0);
        return double.IsFinite(result) ? ComputedValue.Number(result) : ComputedValue.Error(Error.Num);
    }
}
