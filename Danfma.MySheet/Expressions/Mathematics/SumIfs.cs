using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record SumIfs(Expression[] Arguments) : Function
{
    // SUMIFS(sum_range, range1, criteria1, …) — sums sum_range where every (range, criteria) pair matches.
    // The value/criteria ranges are walked as parallel positional cursors (snapshot zero-copy when admitted,
    // else a dense stream) so no per-range List is materialized per evaluation.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (CriteriaScan.CreateWithValue(Arguments, context, out var scan) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var total = 0.0;

        while (scan.MoveNext(out var matched, out var value))
        {
            if (matched && value.TryGetNumber(out var number))
            {
                total += number;
            }
        }

        return ComputedValue.Number(total);
    }
}
