namespace Danfma.MySheet.Expressions;

/// <summary>
/// Shared primitives for the periodic time-value-of-money functions (PMT, PV, FV, NPER, …).
/// All are rearrangements of the annuity equation, with <c>type ∈ {0 = end, 1 = beginning}</c>:
/// <code>
/// r ≠ 0:  pv·(1+r)^n + pmt·(1 + r·type)·((1+r)^n − 1)/r + fv = 0
/// r = 0:  pv + pmt·n + fv = 0
/// </code>
/// These are pure math: degenerate inputs (e.g. nper = 0) yield non-finite results that callers
/// map to the appropriate <see cref="ErrorValue"/>.
/// </summary>
internal static class TimeValueOfMoney
{
    /// <summary>Periodic payment, solving the annuity equation for <c>pmt</c>.</summary>
    public static double Pmt(double rate, double nper, double pv, double fv, double type)
    {
        if (rate == 0)
        {
            return -(pv + fv) / nper;
        }

        var factor = Math.Pow(1 + rate, nper);
        return -(pv * factor + fv) / ((1 + rate * type) * (factor - 1) / rate);
    }

    /// <summary>Present value, solving the annuity equation for <c>pv</c>.</summary>
    public static double Pv(double rate, double nper, double pmt, double fv, double type)
    {
        if (rate == 0)
        {
            return -(fv + pmt * nper);
        }

        var factor = Math.Pow(1 + rate, nper);
        return -(fv + pmt * (1 + rate * type) * (factor - 1) / rate) / factor;
    }

    /// <summary>Future value, solving the annuity equation for <c>fv</c>.</summary>
    public static double Fv(double rate, double nper, double pmt, double pv, double type)
    {
        if (rate == 0)
        {
            return -(pv + pmt * nper);
        }

        var factor = Math.Pow(1 + rate, nper);
        return -(pv * factor + pmt * (1 + rate * type) * (factor - 1) / rate);
    }

    /// <summary>
    /// Number of periods, solving the annuity equation for <c>nper</c>. An impossible amortization
    /// (the logarithm's argument is non-positive) yields <c>NaN</c> for the caller to map to #NUM!.
    /// </summary>
    public static double Nper(double rate, double pmt, double pv, double fv, double type)
    {
        if (rate == 0)
        {
            return -(pv + fv) / pmt;
        }

        var k = pmt * (1 + rate * type) / rate;
        return Math.Log((k - fv) / (k + pv)) / Math.Log(1 + rate);
    }

    /// <summary>
    /// Interest portion of the payment for period <paramref name="per"/> (1-based). The companion
    /// principal portion is <c>Pmt(...) − IPmt(...)</c>. Mirrors the LibreOffice/Apache-POI reference,
    /// which matches Excel. Callers must guard <c>1 ≤ per ≤ nper</c>.
    /// </summary>
    public static double IPmt(
        double rate,
        double per,
        double nper,
        double pv,
        double fv,
        double type
    )
    {
        var pmt = Pmt(rate, nper, pv, fv, type);

        double balanceBefore;
        if (per == 1)
        {
            balanceBefore = type == 1 ? 0 : -pv;
        }
        else
        {
            balanceBefore =
                type == 1 ? Fv(rate, per - 2, pmt, pv, 1) - pmt : Fv(rate, per - 1, pmt, pv, 0);
        }

        return balanceBefore * rate;
    }

    /// <summary>
    /// Robust root finder shared by RATE and IRR. Local methods (Newton/secant) overshoot on the
    /// stiff annuity equation — e.g. a 360-period loan, where <c>(1+guess)^360</c> dwarfs the root —
    /// so this brackets a sign change around <paramref name="guess"/> and bisects it. Bisection
    /// converges on the <em>rate</em> interval (to 1e-7), independent of the residual's currency-scale
    /// magnitude. Returns <c>NaN</c> when no bracket exists (no sign change), which callers map to #NUM!.
    /// </summary>
    public static double Solve(Func<double, double> f, double guess)
    {
        const double tolerance = 1e-7;
        var center = double.IsFinite(guess) ? guess : 0.1;

        // Bracket a sign change by expanding an interval around the guess. Rates live in (−1, ∞), so
        // the lower bound is clamped just above −1.
        double a = center,
            b = center,
            fa = 0,
            fb = 0;
        var spread = Math.Max(Math.Abs(center), 0.1);
        var bracketed = false;

        for (var i = 0; i < 100; i++)
        {
            a = Math.Max(center - spread, -1 + 1e-9);
            b = center + spread;
            fa = f(a);
            fb = f(b);

            if (fa == 0)
            {
                return a;
            }

            if (fb == 0)
            {
                return b;
            }

            if (double.IsFinite(fa) && double.IsFinite(fb) && Math.Sign(fa) != Math.Sign(fb))
            {
                bracketed = true;
                break;
            }

            spread *= 1.6;
        }

        if (!bracketed)
        {
            return double.NaN;
        }

        for (var i = 0; i < 200; i++)
        {
            var mid = (a + b) / 2;
            var fmid = f(mid);

            if (fmid == 0 || (b - a) / 2 < tolerance)
            {
                return mid;
            }

            if (!double.IsFinite(fmid))
            {
                return double.NaN;
            }

            if (Math.Sign(fmid) == Math.Sign(fa))
            {
                a = mid;
                fa = fmid;
            }
            else
            {
                b = mid;
                fb = fmid;
            }
        }

        return (a + b) / 2;
    }
}
