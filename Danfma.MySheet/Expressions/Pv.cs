using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record Pv(Expression[] Arguments) : Function
{
    // PV(rate, nper, pmt, [fv], [type]) — the present value of an annuity.
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

        var result = TimeValueOfMoney.Pv(rate, nper, pmt, fv, type != 0 ? 1 : 0);
        return double.IsFinite(result) ? ComputedValue.Number(result) : ComputedValue.Error(Error.Num);
    }

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();
}
