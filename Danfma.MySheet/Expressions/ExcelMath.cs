using System.Globalization;

namespace Danfma.MySheet.Expressions;

/// <summary>
/// Shared numeric semantics for the scalar math functions. <see cref="Snap"/> mirrors Excel's
/// 15-significant-digit cosmetic rounding on intermediate quotients so decimal-intuitive results
/// survive IEEE-754 noise (1.3/0.2 is 6.499999999999999 in binary, yet Excel documents
/// MROUND(1.3,0.2)=1.4, which requires treating the quotient as 6.5). Factorials/binomials overflow
/// to +infinity, which callers translate to <c>#NUM!</c>.
/// </summary>
internal static class ExcelMath
{
    /// <summary>2^53 — Excel's documented integer-precision bound (GCD/LCM/BASE/DECIMAL).</summary>
    public const double MaxSafeInteger = 9007199254740992d;

    /// <summary>Rounds to 14 significant digits (round-trip through "G14").</summary>
    public static double Snap(double value) =>
        double.IsFinite(value) && value != 0
            ? double.Parse(value.ToString("G14", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
            : value;

    /// <summary>Truncation toward zero at a digit position (TRUNC/ROUNDDOWN semantics).</summary>
    public static double TruncateToDigits(double number, double digits)
    {
        var factor = Math.Pow(10, Math.Truncate(digits));

        return Math.Truncate(Snap(number * factor)) / factor;
    }

    /// <summary>n! for a non-negative integer; +infinity above 170 (caller maps to #NUM!).</summary>
    public static double Factorial(double count)
    {
        var result = 1d;

        for (var i = 2d; i <= count; i++)
        {
            result *= i;
        }

        return result;
    }

    /// <summary>Binomial coefficient C(total, chosen) for integers with 0 &lt;= chosen &lt;= total,
    /// computed multiplicatively and rounded (the exact result is an integer).</summary>
    public static double Binomial(double total, double chosen)
    {
        var k = Math.Min(chosen, total - chosen);
        var result = 1d;

        for (var i = 1d; i <= k; i++)
        {
            result = result * (total - k + i) / i;
        }

        return Math.Round(result);
    }

    public static long Gcd(long left, long right)
    {
        while (right != 0)
        {
            (left, right) = (right, left % right);
        }

        return left;
    }
}
