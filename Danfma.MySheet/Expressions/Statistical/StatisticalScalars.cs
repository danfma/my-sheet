using MemoryPack;

namespace Danfma.MySheet.Expressions.Statistical;

// The scalar statistical functions: Fisher transform pair, the standard normal DENSITY (PHI —
// note GAUSS needs the normal CDF/erf and is deferred to the distributions phase, F4),
// permutation counts, and PROB over a discrete distribution.

[MemoryPackable]
public sealed partial record Fisher(Expression[] Arguments) : Function
{
    // FISHER(x) — 0.5·ln((1+x)/(1−x)); x ≤ −1 or x ≥ 1 → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var x) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return x is <= -1 or >= 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(0.5 * Math.Log((1 + x) / (1 - x)));
    }
}

[MemoryPackable]
public sealed partial record FisherInv(Expression[] Arguments) : Function
{
    // FISHERINV(y) — (e^(2y) − 1)/(e^(2y) + 1), the inverse of FISHER.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var y) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Tanh(y));
    }
}

[MemoryPackable]
public sealed partial record Phi(Expression[] Arguments) : Function
{
    // PHI(x) — density of the standard normal distribution: exp(−x²/2)/√(2π).
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var x) is { } error)
        {
            return ComputedValue.Error(error);
        }

        return ComputedValue.Number(Math.Exp(-x * x / 2) / Math.Sqrt(2 * Math.PI));
    }
}

[MemoryPackable]
public sealed partial record Permut(Expression[] Arguments) : Function
{
    // PERMUT(number, number_chosen) — n!/(n−k)!, both truncated; n ≤ 0, k < 0 or n < k → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var n) is { } nError)
        {
            return ComputedValue.Error(nError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var k) is { } kError)
        {
            return ComputedValue.Error(kError);
        }

        n = Math.Truncate(n);
        k = Math.Truncate(k);

        if (n <= 0 || k < 0 || n < k)
        {
            return ComputedValue.Error(Error.Num);
        }

        var result = 1.0;

        for (var i = 0.0; i < k; i++)
        {
            result *= n - i;
        }

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record PermutationA(Expression[] Arguments) : Function
{
    // PERMUTATIONA(number, number_chosen) — permutations WITH repetition: n^k, both truncated;
    // either negative → #NUM!.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        if (Arguments[0].Evaluate(context).CoerceToNumber(out var n) is { } nError)
        {
            return ComputedValue.Error(nError);
        }

        if (Arguments[1].Evaluate(context).CoerceToNumber(out var k) is { } kError)
        {
            return ComputedValue.Error(kError);
        }

        n = Math.Truncate(n);
        k = Math.Truncate(k);

        if (n < 0 || k < 0)
        {
            return ComputedValue.Error(Error.Num);
        }

        var result = Math.Pow(n, k);

        return double.IsFinite(result)
            ? ComputedValue.Number(result)
            : ComputedValue.Error(Error.Num);
    }
}

[MemoryPackable]
public sealed partial record Prob(Expression[] Arguments) : Function
{
    // PROB(x_range, prob_range, lower_limit, [upper_limit]) — sums the probabilities of the x
    // values inside [lower, upper] (upper omitted → exactly lower). Any probability outside
    // (0, 1] or a probability total ≠ 1 → #NUM!; ranges of different lengths → #N/A.
    public override ComputedValue Evaluate(EvaluationContext context)
    {
        // A reference to a missing sheet is a structural #REF! that short-circuits before the ranges are read.
        if (ReferenceGuard.MissingSheet(Arguments, context) is { } missing)
        {
            return ComputedValue.Error(missing);
        }

        var xs = ArgumentFlattening.ExpandComputedValues(Arguments[0], context);
        var probs = ArgumentFlattening.ExpandComputedValues(Arguments[1], context);

        if (xs.Count != probs.Count)
        {
            return ComputedValue.Error(Error.NA);
        }

        if (Arguments[2].Evaluate(context).CoerceToNumber(out var lower) is { } lowerError)
        {
            return ComputedValue.Error(lowerError);
        }

        var upper = lower;

        if (Arguments.Length == 4)
        {
            if (Arguments[3].Evaluate(context).CoerceToNumber(out upper) is { } upperError)
            {
                return ComputedValue.Error(upperError);
            }
        }

        var probTotal = 0.0;
        var result = 0.0;

        for (var i = 0; i < xs.Count; i++)
        {
            if (xs[i].TryGetError(out var cellError) || probs[i].TryGetError(out cellError))
            {
                return ComputedValue.Error(cellError);
            }

            if (!xs[i].TryGetNumber(out var x) || !probs[i].TryGetNumber(out var probability))
            {
                // A non-numeric probability can never total 1 with the rest.
                return ComputedValue.Error(Error.Num);
            }

            if (probability is <= 0 or > 1)
            {
                return ComputedValue.Error(Error.Num);
            }

            probTotal += probability;

            if (x >= lower && x <= upper)
            {
                result += probability;
            }
        }

        // Snap absorbs IEEE-754 noise in the total (0.2+0.3+0.1+0.4 style sums).
        return ExcelMath.Snap(probTotal) != 1
            ? ComputedValue.Error(Error.Num)
            : ComputedValue.Number(result);
    }
}
