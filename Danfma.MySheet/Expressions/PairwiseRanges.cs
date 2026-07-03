namespace Danfma.MySheet.Expressions;

/// <summary>How a pairwise consumer treats a position where either side is non-numeric.</summary>
internal enum PairwisePolicy
{
    /// <summary>Drop the whole pair (bivariate statistics, SUMX2MY2/SUMX2PY2/SUMXMY2).</summary>
    IgnorePair,

    /// <summary>Keep the pair with the non-numeric side as 0 (SUMPRODUCT semantics).</summary>
    TreatAsZero,
}

/// <summary>
/// Expands two range-like arguments positionally into numeric pairs, the shared engine of the
/// bivariate statistics (CORREL, SLOPE, …) and the SUMX* family. Cell errors from either side
/// propagate first; a length mismatch returns the caller's shape error (Excel uses <c>#N/A</c> for
/// the statistics and <c>#VALUE!</c> for SUMPRODUCT).
/// </summary>
internal static class PairwiseRanges
{
    public static Error? Expand(
        Expression xArgument,
        Expression yArgument,
        EvaluationContext context,
        Error shapeMismatchError,
        PairwisePolicy policy,
        out List<(double X, double Y)> pairs
    )
    {
        pairs = [];

        // A reference argument to a missing sheet is a STRUCTURAL #REF! that short-circuits the whole
        // pairwise computation (the SUMX* family and every bivariate statistic), before enumeration — ahead
        // of the caller's shape/value policy.
        if (ReferenceGuard.MissingSheet(xArgument, context) is { } xMissing)
        {
            return xMissing;
        }

        if (ReferenceGuard.MissingSheet(yArgument, context) is { } yMissing)
        {
            return yMissing;
        }

        var xs = ArgumentFlattening.ExpandComputedValues(xArgument, context);
        var ys = ArgumentFlattening.ExpandComputedValues(yArgument, context);

        if (xs.Count != ys.Count)
        {
            return shapeMismatchError;
        }

        for (var i = 0; i < xs.Count; i++)
        {
            if (xs[i].TryGetError(out var error) || ys[i].TryGetError(out error))
            {
                return error;
            }

            var xIsNumber = xs[i].TryGetNumber(out var x);
            var yIsNumber = ys[i].TryGetNumber(out var y);

            if (xIsNumber && yIsNumber)
            {
                pairs.Add((x, y));
            }
            else if (policy == PairwisePolicy.TreatAsZero)
            {
                pairs.Add((xIsNumber ? x : 0, yIsNumber ? y : 0));
            }
        }

        return null;
    }
}
