using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record SumIfs(Expression[] Arguments) : Function
{
    // SUMIFS(sum_range, range1, criteria1, …) — sums sum_range where every (range, criteria) pair matches.
    public override object? Compute(EvaluationContext context)
    {
        var sumRange = ArgumentFlattening.Expand(Arguments[0], context);
        var ranges = new List<List<object?>>();
        var criterias = new List<Criteria>();

        for (var i = 1; i + 1 < Arguments.Length; i += 2)
        {
            ranges.Add(ArgumentFlattening.Expand(Arguments[i], context));
            criterias.Add(Criteria.Parse(Arguments[i + 1].Compute(context)));
        }

        var length = sumRange.Count;

        foreach (var range in ranges)
        {
            if (range.Count != length)
            {
                return ErrorValue.NotValue;
            }
        }

        var total = 0.0;

        for (var i = 0; i < length; i++)
        {
            if (AllMatch(ranges, criterias, i) && sumRange[i] is double number)
            {
                total += number;
            }
        }

        return total;
    }

    private static bool AllMatch(List<List<object?>> ranges, List<Criteria> criterias, int index)
    {
        for (var j = 0; j < criterias.Count; j++)
        {
            if (!criterias[j].Matches(ranges[j][index]))
            {
                return false;
            }
        }

        return true;
    }
}
