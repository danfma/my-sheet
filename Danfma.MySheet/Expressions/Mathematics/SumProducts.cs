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
        var arrays = new List<List<ComputedValue>>(Arguments.Length);

        foreach (var argument in Arguments)
        {
            arrays.Add(ArgumentFlattening.ExpandComputedValues(argument, context));
        }

        var length = arrays[0].Count;

        foreach (var array in arrays)
        {
            if (array.Count != length)
            {
                return ComputedValue.Error(Error.Value);
            }
        }

        var total = 0.0;

        for (var i = 0; i < length; i++)
        {
            var product = 1.0;

            foreach (var array in arrays)
            {
                if (array[i].TryGetError(out var cellError))
                {
                    return ComputedValue.Error(cellError);
                }

                // Documented rule: array entries that are not numeric are treated as zero.
                product *= array[i].TryGetNumber(out var number) ? number : 0;
            }

            total += product;
        }

        return ComputedValue.Number(total);
    }
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
        if (PairwiseRanges.Expand(
                arguments[0], arguments[1], context, Error.NA, PairwisePolicy.IgnorePair,
                out var pairs)
            is { } error)
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
