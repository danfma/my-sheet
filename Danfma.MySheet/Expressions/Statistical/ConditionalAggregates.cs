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
        // A range (or average_range) over a missing sheet is a structural #REF!.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));
        var range = ArgumentFlattening.ExpandCached(Arguments[0], context, out var snapshot);

        // Single-range numeric `=k` criterion → O(1) via the numeric-equality map (Σ/count of the matching
        // cells). Any other shape scans the cached snapshot(s) linearly (no cell re-read).
        if (
            Arguments.Length < 3
            && snapshot is not null
            && criteria.TryGetNumericEquality(out var key)
        )
        {
            var (keySum, keyCount) = snapshot.NumericEquality(key);
            return keyCount == 0
                ? ComputedValue.Error(Error.DivZero)
                : ComputedValue.Number(keySum / keyCount);
        }

        var averageRange =
            Arguments.Length == 3
                ? ArgumentFlattening.ExpandCached(Arguments[2], context, out _)
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

        return count == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(total / count);
    }
}

[MemoryPackable]
public sealed partial record AverageIfs(Expression[] Arguments) : Function
{
    // AVERAGEIFS(average_range, criteria_range1, criteria1, …) — averages where every pair matches.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (CriteriaScan.CreateWithValue(Arguments, context, out var scan) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var total = 0.0;
        var count = 0;

        while (scan.MoveNext(out var matched, out var value))
        {
            if (matched && value.TryGetNumber(out var number))
            {
                total += number;
                count++;
            }
        }

        return count == 0
            ? ComputedValue.Error(Error.DivZero)
            : ComputedValue.Number(total / count);
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
/// The shared MAXIFS/MINIFS reducer: scans the value/criteria ranges as parallel positional cursors
/// (<see cref="CriteriaScan"/>) and tracks the extreme matching numeric cell — no per-range list.
/// </summary>
file static class CriteriaPairs
{
    public static ComputedValue Extreme(
        Expression[] arguments,
        EvaluationContext context,
        bool larger
    )
    {
        if (CriteriaScan.CreateWithValue(arguments, context, out var scan) is { } error)
        {
            return ComputedValue.Error(error);
        }

        var found = false;
        var extreme = 0.0;

        while (scan.MoveNext(out var matched, out var value))
        {
            if (!matched || !value.TryGetNumber(out var number))
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
