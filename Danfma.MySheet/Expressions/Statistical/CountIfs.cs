using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountIfs(Expression[] Arguments) : Function
{
    // COUNTIFS(range1, criteria1, …) — counts positions where every (range, criteria) pair matches. The
    // criteria ranges are walked as parallel positional cursors (snapshot zero-copy when admitted, else a
    // dense stream) so no per-range List is materialized per evaluation.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (CriteriaScan.CreateCountOnly(Arguments, context, out var scan) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var count = 0;

        while (scan.MoveNext(out var matched, out _))
        {
            if (matched)
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
