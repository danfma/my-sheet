using MemoryPack;

namespace Danfma.MySheet.Expressions.Mathematics;

// SUMPRODUCT and the SUMX* family. Note the asymmetry, straight from the docs: SUMPRODUCT treats
// non-numeric entries as ZERO (a text cell zeroes only its own factor) and reports shape
// mismatches as #VALUE!; the SUMX* functions DROP the whole pair when either side is non-numeric
// and report length mismatches as #N/A (PairwiseRanges.IgnorePair).

[MemoryPackable]
public sealed partial record SumProduct(Expression[] Arguments) : Function
{
    // SUMPRODUCT(array1, [array2], …) — Σ over positions of the product across arrays.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF! that short-circuits before any array is read.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        // Walk the arrays as parallel positional cursors (PositionalRange: snapshot zero-copy → dense
        // rectangle stream → materialized fallback) instead of materializing one List per argument. The
        // only allocation is this tiny per-argument cursor array; the O(cells) range data is never copied.
        // OpenArrayOrRange is the SUMPRODUCT-only factory: it adds the element-wise array backing, so a
        // COMPUTED argument — (A1:A3<>0)*1, ROW(A1:A3), IF(A1:A3>0,1,0) — is a vector rather than a single
        // collapsed scalar. The *IFS family keeps plain Open, where Excel requires real ranges.
        var ranges = new PositionalRange[Arguments.Length];
        ranges[0] = PositionalRange.OpenArrayOrRange(Arguments[0], context);
        var length = ranges[0].Count;

        // Every argument's cell count AND shape is known up front, so validate all dimensions before scanning
        // any value — a mismatch is #VALUE! ahead of any cell error, exactly like the pre-refactor code.
        for (var a = 1; a < ranges.Length; a++)
        {
            ranges[a] = PositionalRange.OpenArrayOrRange(Arguments[a], context);

            if (ranges[a].Count != length || !SameShape(ranges[0], ranges[a]))
            {
                return ComputedValue.Error(Error.Value);
            }
        }

        var total = 0.0;

        for (var i = 0; i < length; i++)
        {
            var product = 1.0;

            // Position-major, then array-major: the first cell error in that order wins, matching the old
            // scan order bit for bit.
            for (var a = 0; a < ranges.Length; a++)
            {
                var cell = ranges[a].Next();

                if (cell.TryGetError(out var cellError))
                {
                    return ComputedValue.Error(cellError);
                }

                // Documented rule: array entries that are not numeric are treated as zero.
                product *= cell.TryGetNumber(out var number) ? number : 0;
            }

            total += product;
        }

        return ComputedValue.Number(total);
    }

    // The documented rule is about DIMENSIONS, not the cell count: "The array arguments must have the same
    // dimensions. If they do not, SUMPRODUCT returns the #VALUE! error value." A 3x1 column and a 1x3 row
    // hold the same three cells and are still #VALUE! in Excel, so the count check alone is not enough.
    // A shape of 0 rows means the argument HAS no rectangle (an open range, a union, a defined name or a
    // scalar served by the materialized fallback): the shape is unknown there, so the pair is judged by its
    // count alone rather than rejected — SUMPRODUCT(MyName,(A1:A3<>0)*1) is legal in Excel.
    private static bool SameShape(in PositionalRange first, in PositionalRange other) =>
        first.Rows == 0
        || other.Rows == 0
        || (first.Rows == other.Rows && first.Columns == other.Columns);
}

[MemoryPackable]
public sealed partial record SumX2MY2(Expression[] Arguments) : Function
{
    // SUMX2MY2(array_x, array_y) — Σ(x² − y²).
    public override ComputedValue Evaluate(EvaluationContext context) =>
        SumOfPairs.Compute(Arguments, context, static (x, y) => x * x - y * y);
}

[MemoryPackable]
public sealed partial record SumX2PY2(Expression[] Arguments) : Function
{
    // SUMX2PY2(array_x, array_y) — Σ(x² + y²).
    public override ComputedValue Evaluate(EvaluationContext context) =>
        SumOfPairs.Compute(Arguments, context, static (x, y) => x * x + y * y);
}

[MemoryPackable]
public sealed partial record SumXMY2(Expression[] Arguments) : Function
{
    // SUMXMY2(array_x, array_y) — Σ(x − y)².
    public override ComputedValue Evaluate(EvaluationContext context) =>
        SumOfPairs.Compute(Arguments, context, static (x, y) => (x - y) * (x - y));
}

file static class SumOfPairs
{
    public static ComputedValue Compute(
        Expression[] arguments,
        EvaluationContext context,
        Func<double, double, double> term
    )
    {
        if (
            PairwiseRanges.Expand(
                arguments[0],
                arguments[1],
                context,
                Error.NA,
                PairwisePolicy.IgnorePair,
                out var pairs
            ) is
            { } error
        )
        {
            return ComputedValue.Error(error);
        }

        var total = 0.0;

        foreach (var (x, y) in pairs)
        {
            total += term(x, y);
        }

        return ComputedValue.Number(total);
    }
}
