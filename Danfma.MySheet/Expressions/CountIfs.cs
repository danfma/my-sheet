using MemoryPack;

namespace Danfma.MySheet.Expressions;

[MemoryPackable]
public sealed partial record CountIfs(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var ranges = new List<List<object?>>();
        var criterias = new List<Criteria>();

        for (var i = 0; i + 1 < Arguments.Length; i += 2)
        {
            ranges.Add(ArgumentFlattening.Expand(Arguments[i], context));
            criterias.Add(Criteria.Parse(Arguments[i + 1].Compute(context)));
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

    public override object? Compute(EvaluationContext context) => Evaluate(context).AsObject();

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
