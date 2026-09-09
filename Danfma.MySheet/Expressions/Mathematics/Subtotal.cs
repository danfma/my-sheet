using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record Subtotal(Expression[] Arguments) : Function
{
    // SUBTOTAL(function_num, ref1, [ref2], …) — applies the aggregate selected by function_num
    // (1-11; 101-111 behave identically here: MySheet has no hidden-row model, a documented model
    // limit) while IGNORING any referenced cell whose own formula is a SUBTOTAL node, so stacked
    // subtotal rows are not double counted (the real nested-subtotal rule). Rows excluded by a
    // filter do not apply either (no filter model). An invalid function_num → #VALUE!.
    //
    // The scan, the accumulator and the code map live in AggregateCodes, shared with AGGREGATE. The nested
    // skip stays at NestedSkip.Subtotal: Microsoft's SUBTOTAL page documents only "nested subtotals are
    // ignored", while the "nested SUBTOTAL and AGGREGATE" wording appears solely in AGGREGATE's own options
    // table — and the narrow reading is the MEASURED one (Aspose.Cells 26.6.0, 2026-09-09: over A1=1, A2=2,
    // A3="=AGGREGATE(9,0,A1:A2)"=3, =SUBTOTAL(9,A1:A3) is 6, so the nested AGGREGATE is counted).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var rawCode) is { } codeError)
        {
            return ComputedValue.Error(codeError);
        }

        var code = (int)Math.Truncate(rawCode);

        if (code is >= 101 and <= 111)
        {
            code -= 100; // "ignore hidden rows" variants: same behaviour without a hidden-row model
        }

        if (code is < 1 or > 11)
        {
            return ComputedValue.Error(Error.Value);
        }

        var refs = Arguments[1..];

        // A reference argument to a missing sheet is a structural #REF! that short-circuits SUBTOTAL, before
        // any per-cell scan (which would otherwise index a non-existent sheet, or be swallowed by the
        // COUNT/COUNTA codes).
        if (ReferenceGuard.MissingSheet(refs, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var accumulator = new AggregateCodes.Accumulator(code, ignoreErrors: false);

        foreach (var argument in refs)
        {
            if (
                AggregateCodes.Feed(
                    argument,
                    context,
                    ref accumulator,
                    AggregateCodes.NestedSkip.Subtotal
                ) is
                { } error
            )
            {
                return ComputedValue.Error(error);
            }
        }

        return accumulator.Finish();
    }
}
