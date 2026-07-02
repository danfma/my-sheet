using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountIfs(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var ranges = new List<List<ComputedValue>>();
        var criterias = new List<Criteria>();

        for (var i = 0; i + 1 < Arguments.Length; i += 2)
        {
            ranges.Add(ArgumentFlattening.ExpandComputedValues(Arguments[i], context));
            criterias.Add(Criteria.Parse(Arguments[i + 1].Evaluate(context)));
        }

        var length = ranges[0].Count;

        foreach (var range in ranges)
        {
            if (range.Count != length)
            {
                return ComputedValue.Error(Error.Value);
            }
        }

        var count = 0;

        for (var i = 0; i < length; i++)
        {
            if (AllMatch(ranges, criterias, i))
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
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
