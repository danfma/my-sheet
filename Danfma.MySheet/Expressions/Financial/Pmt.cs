using MemoryPack;

namespace Danfma.MySheet.Expressions.Financial;

[MemoryPackable]
public sealed partial record Pmt(Expression[] Arguments) : Function
{
    // PMT(rate, nper, pv, [fv], [type]) — the periodic payment for a loan/annuity.
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

        var result = TimeValueOfMoney.Pmt(rate, nper, pv, fv, type != 0 ? 1 : 0);
        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}
