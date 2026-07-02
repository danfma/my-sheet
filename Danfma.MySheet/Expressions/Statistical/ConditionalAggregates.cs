using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

// The criteria-driven statistical aggregates, mirrors of SUMIF/SUMIFS (Mathematics): parallel
// ranges are aligned by position, criteria go through the shared Criteria engine (wildcards,
// comparison operators). AVERAGEIF(S) with no matching numeric cell → #DIV/0!; MAXIFS/MINIFS with
// no match → 0 (both per the Microsoft docs).

[MemoryPackable]
public sealed partial record AverageIf(Expression[] Arguments) : Function
{
    // AVERAGEIF(range, criteria, [average_range]) — averages average_range (or range) where range
    // matches the criteria.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));
        var range = ArgumentFlattening.ExpandComputedValues(Arguments[0], context);
        var averageRange = Arguments.Length == 3
            ? ArgumentFlattening.ExpandComputedValues(Arguments[2], context)
            : range;

        var total = 0.0;
        var count = 0;
        var length = Math.Min(range.Count, averageRange.Count);

        for (var i = 0; i < length; i++)
        {
            if (criteria.Matches(range[i]) && averageRange[i].TryGetNumber(out var number))
            {
                total += number;
                count++;
            }
        }

        return count == 0 ? ComputedValue.Error(Error.DivZero) : ComputedValue.Number(total / count);
    }
}

[MemoryPackable]
public sealed partial record AverageIfs(Expression[] Arguments) : Function
{
    // AVERAGEIFS(average_range, criteria_range1, criteria1, …) — averages where every pair matches.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (CriteriaPairs.Expand(Arguments, context, out var valueRange, out var pairs) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var total = 0.0;
        var count = 0;

        for (var i = 0; i < valueRange.Count; i++)
        {
            if (CriteriaPairs.AllMatch(pairs, i) && valueRange[i].TryGetNumber(out var number))
            {
                total += number;
                count++;
            }
        }

        return count == 0 ? ComputedValue.Error(Error.DivZero) : ComputedValue.Number(total / count);
    }
}

[MemoryPackable]
public sealed partial record MaxIfs(Expression[] Arguments) : Function
{
    // MAXIFS(max_range, criteria_range1, criteria1, …) — largest matching value; no match → 0.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        CriteriaPairs.Extreme(Arguments, context, larger: true);
}

[MemoryPackable]
public sealed partial record MinIfs(Expression[] Arguments) : Function
{
    // MINIFS(min_range, criteria_range1, criteria1, …) — smallest matching value; no match → 0.
    public override ComputedValue Evaluate(EvaluationContext context) =>
        CriteriaPairs.Extreme(Arguments, context, larger: false);
}

/// <summary>
/// The shared (criteria_range, criteria) plumbing of the *IFS statistical aggregates: expands the
/// value range plus every pair, enforcing equal lengths (<c>#VALUE!</c> on mismatch, like SUMIFS).
/// </summary>
file static class CriteriaPairs
{
    public static Error? Expand(
        Expression[] arguments,
        EvaluationContext context,
        out List<ComputedValue> valueRange,
        out List<(List<ComputedValue> Range, Criteria Criteria)> pairs
    )
    {
        valueRange = ArgumentFlattening.ExpandComputedValues(arguments[0], context);
        pairs = [];

        for (var i = 1; i + 1 < arguments.Length; i += 2)
        {
            var range = ArgumentFlattening.ExpandComputedValues(arguments[i], context);

            if (range.Count != valueRange.Count)
            {
                return Error.Value;
            }

            pairs.Add((range, Criteria.Parse(arguments[i + 1].Evaluate(context))));
        }

        return null;
    }

    public static bool AllMatch(List<(List<ComputedValue> Range, Criteria Criteria)> pairs, int index)
    {
        foreach (var (range, criteria) in pairs)
        {
            if (!criteria.Matches(range[index]))
            {
                return false;
            }
        }

        return true;
    }

    public static ComputedValue Extreme(
        Expression[] arguments,
        EvaluationContext context,
        bool larger
    )
    {
        if (Expand(arguments, context, out var valueRange, out var pairs) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var found = false;
        var extreme = 0.0;

        for (var i = 0; i < valueRange.Count; i++)
        {
            if (!AllMatch(pairs, i) || !valueRange[i].TryGetNumber(out var number))
            {
                continue;
            }

            if (!found || (larger ? number > extreme : number < extreme))
            {
                extreme = number;
                found = true;
            }
        }

        // Documented MAXIFS/MINIFS behaviour: no cell satisfies the criteria → 0.
        return ComputedValue.Number(found ? extreme : 0);
    }
}
