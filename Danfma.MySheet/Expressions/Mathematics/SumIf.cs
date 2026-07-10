using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

[MemoryPackable]
public sealed partial record SumIf(Expression[] Arguments) : Function
{
    // SUMIF(range, criteria, [sum_range]) — sums sum_range (or range) where range matches the criteria.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A range (or sum_range) over a missing sheet is a structural #REF!, not an empty range summing 0.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var criteria = Criteria.Parse(Arguments[1].Evaluate(context));

        var snapshot = Arguments[0] is Reference reference
            ? context.Workbook.TryGetRangeSnapshot(reference, context)
            : null;

        // Single-range (no separate sum_range) numeric `=k` criterion → O(1) via the numeric-equality map,
        // whose per-key Sum is exactly Σ of the matching cells' own numeric values.
        if (
            Arguments.Length < 3
            && snapshot is not null
            && criteria.TryGetNumericEquality(out var key)
        )
        {
            return ComputedValue.Number(snapshot.NumericEquality(key).Sum);
        }

        // Any other shape: a positional cursor over range (and, with a sum_range, a second parallel cursor)
        // — the admitted snapshot is indexed zero-copy, a non-admitted closed rectangle streams through the
        // dense struct enumerator (no allocation), and only an open range/union/scalar falls back to a
        // materialized list. Mirrors the SUMIFS pair-scan idiom (CriteriaScan/PositionalRange), specialized
        // to SUMIF's single (range, sum_range) pair. Threads the ALREADY-probed `snapshot` through instead of
        // letting Open re-probe it: TryGetRangeSnapshot is the second-use ADMISSION check itself, so a second
        // call here (even for the same range within this same evaluation) would eagerly build the snapshot on
        // what must stay range's first, streaming read.
        var range = PositionalRange.Open(Arguments[0], context, snapshot);

        if (Arguments.Length < 3)
        {
            var singleTotal = 0.0;

            for (var i = 0; i < range.Count; i++)
            {
                var cell = range.Next();

                if (criteria.Matches(cell) && cell.TryGetNumber(out var number))
                {
                    singleTotal += number;
                }
            }

            return ComputedValue.Number(singleTotal);
        }

        var sumRange = PositionalRange.Open(Arguments[2], context);
        var total = 0.0;
        var length = Math.Min(range.Count, sumRange.Count);

        for (var i = 0; i < length; i++)
        {
            var cell = range.Next();
            var sumCell = sumRange.Next();

            if (criteria.Matches(cell) && sumCell.TryGetNumber(out var number))
            {
                total += number;
            }
        }

        return ComputedValue.Number(total);
    }
}
