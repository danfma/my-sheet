using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

[MemoryPackable]
public sealed partial record CountIf(Expression[] Arguments) : Function
{
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A range over a missing sheet is a structural #REF!, not an empty range that matches nothing.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));
        var range = ArgumentFlattening.ExpandCached(Arguments[0], context, out var snapshot);

        // Numeric `=k` criterion over a cached range → O(1) via the numeric-equality map; anything else
        // (text/wildcard/comparison) scans the cached snapshot linearly (no cell re-read).
        if (snapshot is not null && criteria.TryGetNumericEquality(out var key))
        {
            return ComputedValue.Number(snapshot.NumericEquality(key).Count);
        }

        var count = 0;

        foreach (var value in range)
        {
            if (criteria.Matches(value))
            {
                count++;
            }
        }

        return ComputedValue.Number(count);
    }
}
