using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record SumIfs(Expression[] Arguments) : Function
{
    // SUMIFS(sum_range, range1, criteria1, …) — sums sum_range where every (range, criteria) pair matches.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A sum_range or criteria range over a missing sheet is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var sumRange = ArgumentFlattening.ExpandComputedValues(Arguments[0], context);
        var ranges = new List<List<ComputedValue>>();
        var criterias = new List<Criteria>();

        for (var i = 1; i + 1 < Arguments.Length; i += 2)
        {
            ranges.Add(ArgumentFlattening.ExpandComputedValues(Arguments[i], context));
            criterias.Add(Criteria.Parse(Arguments[i + 1].Evaluate(context)));
        }

        var length = sumRange.Count;

        foreach (var range in ranges)
        {
            if (range.Count != length)
            {
                return ComputedValue.Error(Error.Value);
            }
        }

        var total = 0.0;

        for (var i = 0; i < length; i++)
        {
            if (AllMatch(ranges, criterias, i) && sumRange[i].TryGetNumber(out var number))
            {
                total += number;
            }
        }

        return ComputedValue.Number(total);
    }

    private static bool AllMatch(List<List<ComputedValue>> ranges, List<Criteria> criterias, int index)
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
