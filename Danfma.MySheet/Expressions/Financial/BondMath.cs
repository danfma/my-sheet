using System;
using System.Collections.Generic;

namespace Danfma.MySheet.Expressions.Financial;

/// <summary>
/// Day-count and cash-flow math shared by the wave-6 financial functions (bonds, T-bills, depreciation,
/// coupon schedules). It is a faithful port of the conventions Excel itself uses for these functions —
/// cross-checked value-for-value against the <c>ExcelFinancialFunctions</c> oracle by fuzzing.
///
/// It builds on <see cref="DayCount"/> (the wave-5 YEARFRAC helper) for the plain Actual and European
/// 30/360 counts, but the bond world needs two conventions that YEARFRAC never exposes and that differ
/// from <see cref="DayCount.Nasd360Days"/> in month-end edge cases:
/// <list type="bullet">
/// <item>the US (NASD) 30/360 count has a <em>ModifyStartDate</em> vs <em>ModifyBothDates</em> flavour
/// (COUPDAYS/accrual use the latter);</item>
/// <item>actual/actual bonds measure the year as the average length of the calendar years the coupon
/// period straddles, not the plain average YEARFRAC uses.</item>
/// </list>
/// Because of that, the US 30/360 count, the actual/actual "days in year", and the backward coupon
/// schedule live here rather than being forced onto the YEARFRAC helper. Bases follow Excel: 0 = US 30/360,
/// 1 = actual/actual, 2 = actual/360, 3 = actual/365, 4 = European 30/360.
/// </summary>
internal static class BondMath
{
    internal enum Method360
    {
        ModifyStartDate,
        ModifyBothDates,
    }

    // Actual calendar days between two dates (order: after - before), reusing the wave-5 helper.
    internal static int Days(DateTime after, DateTime before) => DayCount.ActualDays(before, after);

    private static bool IsLastDayOfMonth(DateTime d) =>
        d.Day == DateTime.DaysInMonth(d.Year, d.Month);

    private static bool LastFeb(DateTime d) => d.Month == 2 && IsLastDayOfMonth(d);

    /// <summary>US (NASD) 30/360 count with Excel's start/both-date month-end rules.</summary>
    internal static int Diff360Us(DateTime s, DateTime e, Method360 method)
    {
        int sd = s.Day,
            sm = s.Month,
            sy = s.Year,
            ed = e.Day,
            em = e.Month,
            ey = e.Year;
        var both = method == Method360.ModifyBothDates;

        if (LastFeb(e) && (LastFeb(s) || both))
        {
            ed = 30;
        }

        if (ed == 31 && (sd >= 30 || both))
        {
            ed = 30;
        }

        if (sd == 31)
        {
            sd = 30;
        }

        if (LastFeb(s))
        {
            sd = 30;
        }

        return (ey - sy) * 360 + (em - sm) * 30 + (ed - sd);
    }

    private static int Diff365(DateTime s, DateTime e)
    {
        int sd = s.Day,
            ed = e.Day;

        if (sd > 28 && s.Month == 2)
        {
            sd = 28;
        }

        if (ed > 28 && e.Month == 2)
        {
            ed = 28;
        }

        var startd = new DateTime(s.Year, s.Month, sd);
        var endd = new DateTime(e.Year, e.Month, ed);
        return (e.Year - s.Year) * 365 + Days(endd, startd);
    }

    /// <summary>Shift a date by whole months, optionally snapping to the last day of the target month.</summary>
    internal static DateTime ChangeMonth(DateTime d, int numMonths, bool returnLastDay)
    {
        var t = d.AddMonths(numMonths);
        return returnLastDay
            ? new DateTime(t.Year, t.Month, DateTime.DaysInMonth(t.Year, t.Month))
            : t;
    }

    private static (DateTime front, DateTime trailing, double acc) DatesAggregate1(
        DateTime startDate,
        DateTime endDate,
        int numMonths,
        bool returnLastMonth,
        Func<DateTime, DateTime, double> f,
        double acc
    )
    {
        var front = startDate;
        var trailing = startDate;

        while (true)
        {
            var stop = numMonths > 0 ? front >= endDate : front <= endDate;

            if (stop)
            {
                return (front, trailing, acc);
            }

            trailing = front;
            front = ChangeMonth(front, numMonths, returnLastMonth);
            acc += f(front, trailing);
        }
    }

    private static (DateTime pcd, DateTime ncd) FindPcdNcd(
        DateTime startDate,
        DateTime endDate,
        int numMonths,
        bool returnLastMonth
    )
    {
        var (p, n, _) = DatesAggregate1(
            startDate,
            endDate,
            numMonths,
            returnLastMonth,
            (_, _) => 0d,
            0d
        );
        return (p, n);
    }

    private static int Freq2Months(double freq) => 12 / (int)freq;

    private static (DateTime pcd, DateTime ncd) FindCouponDates(
        DateTime settl,
        DateTime mat,
        double freq
    )
    {
        var endMonth = IsLastDayOfMonth(mat);
        var numMonths = -Freq2Months(freq);
        return FindPcdNcd(mat, settl, numMonths, endMonth);
    }

    internal static DateTime CoupPcd(DateTime settl, DateTime mat, double freq) =>
        FindCouponDates(settl, mat, freq).pcd;

    internal static DateTime CoupNcd(DateTime settl, DateTime mat, double freq) =>
        FindCouponDates(settl, mat, freq).ncd;

    internal static double CoupNum(DateTime settl, DateTime mat, double freq)
    {
        var pcd = CoupPcd(settl, mat, freq);
        double months = (mat.Year - pcd.Year) * 12 + (mat.Month - pcd.Month);
        return months * freq / 12d;
    }

    internal static double CoupDays(int basis, DateTime settl, DateTime mat, double freq)
    {
        if (basis == 1)
        {
            var (p, n) = FindCouponDates(settl, mat, freq);
            return Days(n, p);
        }

        return basis == 3 ? 365d / freq : 360d / freq;
    }

    internal static double CoupDaysBs(int basis, DateTime settl, DateTime mat, double freq)
    {
        var pcd = CoupPcd(settl, mat, freq);
        return basis switch
        {
            0 => Diff360Us(pcd, settl, Method360.ModifyStartDate),
            4 => DayCount.Euro360Days(pcd, settl),
            _ => Days(settl, pcd),
        };
    }

    internal static double CoupDaysNc(int basis, DateTime settl, DateTime mat, double freq)
    {
        var pcd = CoupPcd(settl, mat, freq);
        var ncd = CoupNcd(settl, mat, freq);
        return basis switch
        {
            0 => Diff360Us(pcd, ncd, Method360.ModifyBothDates)
                - Diff360Us(pcd, settl, Method360.ModifyStartDate),
            4 => DayCount.Euro360Days(settl, ncd),
            _ => Days(ncd, settl),
        };
    }

    /// <summary>
    /// The yield-independent coupon-schedule facts <see cref="Price"/>, <see cref="Yield"/> and
    /// <see cref="Duration"/> need, computed with a single backward walk (<see cref="FindCouponDates"/>)
    /// instead of the 2-3 independent walks that calling <see cref="CoupNum"/>/<see cref="CoupPcd"/>/
    /// <see cref="CoupDays"/> separately would each perform. Reused across every iteration of the YIELD
    /// solver, where only the yield itself varies.
    /// </summary>
    private readonly record struct CouponSchedule(DateTime Pcd, DateTime Ncd, double N, double E);

    private static CouponSchedule GetCouponSchedule(
        DateTime settl,
        DateTime mat,
        double freq,
        int basis
    )
    {
        var (pcd, ncd) = FindCouponDates(settl, mat, freq);
        double months = (mat.Year - pcd.Year) * 12 + (mat.Month - pcd.Month);
        var n = months * freq / 12d;
        var e =
            basis == 1 ? Days(ncd, pcd)
            : basis == 3 ? 365d / freq
            : 360d / freq;
        return new CouponSchedule(pcd, ncd, n, e);
    }

    /// <summary>Basis-specific day count between two dates (Numerator = accrual-style actual for 2/3).</summary>
    internal static double DaysBetween(int basis, DateTime a, DateTime b, bool numerator) =>
        basis switch
        {
            0 => Diff360Us(a, b, Method360.ModifyStartDate),
            4 => DayCount.Euro360Days(a, b),
            2 => numerator ? Days(b, a) : Diff360Us(a, b, Method360.ModifyStartDate),
            3 => numerator ? Days(b, a) : Diff365(a, b),
            _ => Days(b, a),
        };

    internal static double DaysInYear(int basis, DateTime issue, DateTime settl)
    {
        if (basis == 3)
        {
            return 365d;
        }

        if (basis != 1)
        {
            return 360d;
        }

        if (!LessOrEqualToAYearApart(issue, settl))
        {
            var totYears = settl.Year - issue.Year + 1;
            var totDays = Days(new DateTime(settl.Year + 1, 1, 1), new DateTime(issue.Year, 1, 1));
            return (double)totDays / totYears;
        }

        return ConsiderAsBisestile(issue, settl) ? 366d : 365d;
    }

    private static bool LessOrEqualToAYearApart(DateTime d1, DateTime d2) =>
        d1.Year == d2.Year
        || (
            d2.Year == d1.Year + 1
            && (d1.Month > d2.Month || (d1.Month == d2.Month && d1.Day >= d2.Day))
        );

    private static bool IsFeb29Between(DateTime d1, DateTime d2)
    {
        if (d1.Year == d2.Year && DateTime.IsLeapYear(d1.Year))
        {
            return d1.Month <= 2 && d2.Month > 2;
        }

        if (d1.Year == d2.Year)
        {
            return false;
        }

        if (d2.Year == d1.Year + 1)
        {
            if (DateTime.IsLeapYear(d1.Year))
            {
                return d1.Month <= 2;
            }

            return DateTime.IsLeapYear(d2.Year) && d2.Month > 2;
        }

        return false;
    }

    private static bool ConsiderAsBisestile(DateTime d1, DateTime d2) =>
        (d1.Year == d2.Year && DateTime.IsLeapYear(d1.Year))
        || (d2.Month == 2 && d2.Day == 29)
        || IsFeb29Between(d1, d2);

    // ---- bond price / yield / duration ----

    internal static double Price(
        DateTime settl,
        DateTime mat,
        double rate,
        double yld,
        double redemption,
        double freq,
        int basis
    ) =>
        Price(
            GetCouponSchedule(settl, mat, freq, basis),
            settl,
            rate,
            yld,
            redemption,
            freq,
            basis
        );

    // Same arithmetic as the public overload, but takes an already-computed schedule so the YIELD solver
    // (which calls this once per iteration, only varying `yld`) doesn't re-walk the coupon schedule every time.
    private static double Price(
        CouponSchedule schedule,
        DateTime settl,
        double rate,
        double yld,
        double redemption,
        double freq,
        int basis
    )
    {
        var n = schedule.N;
        var a = DaysBetween(basis, schedule.Pcd, settl, true);
        var e = schedule.E;
        var dsc = e - a;
        var coupon = 100d * rate / freq;
        var accrued = 100d * rate / freq * a / e;

        double Pv(double k) => Math.Pow(1d + yld / freq, k - 1d + dsc / e);

        if (n == 1d)
        {
            return (redemption + coupon) / (1d + dsc / e * yld / freq) - accrued;
        }

        var pv = redemption / Pv(n);

        for (var k = 1; k <= (int)n; k++)
        {
            pv += coupon / Pv(k);
        }

        return pv - accrued;
    }

    internal static double Yield(
        DateTime settl,
        DateTime mat,
        double rate,
        double pr,
        double redemption,
        double freq,
        int basis
    )
    {
        var schedule = GetCouponSchedule(settl, mat, freq, basis);
        var n = schedule.N;
        var a = DaysBetween(basis, schedule.Pcd, settl, true);
        var e = schedule.E;
        var dsr = e - a;

        if (n <= 1d)
        {
            var k = (redemption / 100d + rate / freq) / (pr / 100d + a / e * rate / freq) - 1d;
            return k * freq * e / dsr;
        }

        return TimeValueOfMoney.Solve(
            y => Price(schedule, settl, rate, y, redemption, freq, basis) - pr,
            0.05
        );
    }

    internal static double Duration(
        DateTime settl,
        DateTime mat,
        double coupon,
        double yld,
        double freq,
        int basis,
        bool modified
    )
    {
        var schedule = GetCouponSchedule(settl, mat, freq, basis);
        var dbc = basis switch
        {
            0 => Diff360Us(schedule.Pcd, settl, Method360.ModifyStartDate),
            4 => DayCount.Euro360Days(schedule.Pcd, settl),
            _ => Days(settl, schedule.Pcd),
        };
        var e = schedule.E;
        var n = schedule.N;
        var dsc = e - dbc;
        var x1 = dsc / e;
        var x2 = x1 + n - 1d;
        var x3 = yld / freq + 1d;
        var x4 = Math.Pow(x3, x2);
        var term1 = x2 * 100d / x4;
        var term3 = 100d / x4;
        double numerator = 0,
            denominator = 0;

        for (var index = 1; index <= (int)n; index++)
        {
            var x5 = index - 1d + x1;
            var x6 = Math.Pow(x3, x5);
            var x7 = 100d * coupon / freq / x6;
            numerator += x7 * x5;
            denominator += x7;
        }

        var term5 = term1 + numerator;
        var term6 = term3 + denominator;
        return modified ? (term5 / term6 / freq) / x3 : term5 / term6 / freq;
    }

    // ---- accrued interest ----

    internal static double AccrIntM(
        DateTime issue,
        DateTime settl,
        double rate,
        double par,
        int basis
    ) => par * rate * (DaysBetween(basis, issue, settl, true) / DaysInYear(basis, issue, settl));

    internal static double AccrInt(
        DateTime issue,
        DateTime firstInterest,
        DateTime settl,
        double rate,
        double par,
        double freq,
        int basis,
        int calcMethod
    )
    {
        var numMonths = Freq2Months(freq);
        var numMonthsNeg = -numMonths;
        var endMonthBond = IsLastDayOfMonth(firstInterest);

        var pcd =
            settl > firstInterest && calcMethod == 1
                ? FindPcdNcd(firstInterest, settl, numMonths, endMonthBond).pcd
                : ChangeMonth(firstInterest, numMonthsNeg, endMonthBond);

        var firstDate = issue > pcd ? issue : pcd;
        var days = DaysBetween(basis, firstDate, settl, true);
        var coupDays0 = CoupDays(basis, pcd, firstInterest, freq);

        double Aggregate(DateTime p, DateTime ncd)
        {
            var fd = issue > p ? issue : p;
            double d;

            if (basis == 0)
            {
                var psaMethod = issue > p ? Method360.ModifyStartDate : Method360.ModifyBothDates;
                d = Diff360Us(fd, ncd, psaMethod);
            }
            else
            {
                d = DaysBetween(basis, fd, ncd, true);
            }

            double cd;

            if (basis == 0)
            {
                cd = Diff360Us(p, ncd, Method360.ModifyBothDates);
            }
            else if (basis == 3)
            {
                cd = 365d / freq;
            }
            else
            {
                cd = DaysBetween(basis, p, ncd, false);
            }

            return issue <= p ? calcMethod : d / cd;
        }

        var (_, _, a) = DatesAggregate1(
            pcd,
            issue,
            numMonthsNeg,
            endMonthBond,
            Aggregate,
            days / coupDays0
        );
        return par * rate / freq * a;
    }

    // ---- simple discount / interest bonds ----

    internal static double IntRate(
        DateTime s,
        DateTime m,
        double investment,
        double redemption,
        int basis
    )
    {
        double dim = DaysBetween(basis, s, m, true),
            b = DaysInYear(basis, s, m);
        return (redemption - investment) / investment * b / dim;
    }

    internal static double Received(
        DateTime s,
        DateTime m,
        double investment,
        double discount,
        int basis
    )
    {
        double dim = DaysBetween(basis, s, m, true),
            b = DaysInYear(basis, s, m);
        return investment / (1d - discount * dim / b);
    }

    internal static double Disc(DateTime s, DateTime m, double pr, double redemption, int basis)
    {
        double dim = DaysBetween(basis, s, m, true),
            b = DaysInYear(basis, s, m);
        return (-pr / redemption + 1d) * b / dim;
    }

    internal static double PriceDisc(
        DateTime s,
        DateTime m,
        double discount,
        double redemption,
        int basis
    )
    {
        double dim = DaysBetween(basis, s, m, true),
            b = DaysInYear(basis, s, m);
        return redemption - discount * redemption * dim / b;
    }

    internal static double YieldDisc(
        DateTime s,
        DateTime m,
        double pr,
        double redemption,
        int basis
    )
    {
        double dim = DaysBetween(basis, s, m, true),
            b = DaysInYear(basis, s, m);
        return (redemption - pr) / pr * b / dim;
    }

    internal static double PriceMat(
        DateTime s,
        DateTime m,
        DateTime issue,
        double rate,
        double yld,
        int basis
    )
    {
        double b = DaysInYear(basis, issue, s),
            dim = DaysBetween(basis, issue, m, true),
            a = DaysBetween(basis, issue, s, true),
            dsm = dim - a;
        return (100d + dim / b * rate * 100d) / (1d + dsm / b * yld) - a / b * rate * 100d;
    }

    internal static double YieldMat(
        DateTime s,
        DateTime m,
        DateTime issue,
        double rate,
        double pr,
        int basis
    )
    {
        double b = DaysInYear(basis, issue, s),
            dim = DaysBetween(basis, issue, m, true),
            a = DaysBetween(basis, issue, s, true),
            dsm = dim - a;
        var term1 = dim / b * rate + 1d - pr / 100d - a / b * rate;
        var term2 = pr / 100d + a / b * rate;
        var term3 = b / dsm;
        return term1 / term2 * term3;
    }

    // ---- T-bills ----

    internal static double TBillEq(DateTime s, DateTime m, double discount)
    {
        double dsm = Days(m, s);

        if (dsm > 182d)
        {
            var price = (100d - discount * 100d * dsm / 360d) / 100d;
            var days = dsm == 366d ? 366d : 365d;
            var term2 = Math.Sqrt(
                Math.Pow(dsm / days, 2d) - (2d * dsm / days - 1d) * (1d - 1d / price)
            );
            var term3 = 2d * dsm / days - 1d;
            return 2d * (term2 - dsm / days) / term3;
        }

        return 365d * discount / (360d - discount * dsm);
    }

    internal static double TBillPrice(DateTime s, DateTime m, double discount)
    {
        double dsm = Days(m, s);
        return 100d * (1d - discount * dsm / 360d);
    }

    internal static double TBillYield(DateTime s, DateTime m, double pr)
    {
        double dsm = Days(m, s);
        return (100d - pr) / pr * 360d / dsm;
    }

    // ---- depreciation ----

    internal static double Sln(double cost, double salvage, double life) => (cost - salvage) / life;

    internal static double Syd(double cost, double salvage, double life, double per) =>
        (cost - salvage) * (life - per + 1d) * 2d / (life * (life + 1d));

    private static double DeprRate(double cost, double salvage, double life) =>
        Math.Round(1d - Math.Pow(salvage / cost, 1d / life), 3, MidpointRounding.AwayFromZero);

    internal static double Db(double cost, double salvage, double life, double period, double month)
    {
        var rate = DeprRate(cost, salvage, life);
        double DeprFirst() => cost * rate * month / 12d;
        double DeprLast(double tot) => (cost - tot) * rate * (12d - month) / 12d;
        double DeprPer(double tot) => (cost - tot) * rate;

        double totDepr = 0,
            per = 0;

        while (true)
        {
            var ip = (int)per;

            if (ip == 0)
            {
                var d = DeprFirst();

                if ((int)period <= 1)
                {
                    return d;
                }

                totDepr = d;
                per += 1;
                continue;
            }

            if (ip == (int)period - 1)
            {
                return DeprPer(totDepr);
            }

            if (ip == (int)life - 1)
            {
                return DeprLast(totDepr);
            }

            totDepr += DeprPer(totDepr);
            per += 1;
        }
    }

    private static double Rest(double x) => x - (int)x;

    private static double TotalDepr(
        double cost,
        double salvage,
        double life,
        double period,
        double factor,
        bool straightLine
    )
    {
        double totDepr = 0,
            per = 0;
        var frac = Rest(period);
        double Ddb(double t) => Math.Min((cost - t) * (factor / life), cost - salvage - t);
        double Sl(double t, double aPeriod) => Sln(cost - t, salvage, life - aPeriod);

        while (true)
        {
            double ddbD = Ddb(totDepr),
                slnD = Sl(totDepr, per);
            var isSln = straightLine && ddbD < slnD;
            var depr = isSln ? slnD : ddbD;
            var newTot = totDepr + depr;

            if ((int)period == 0)
            {
                return newTot * frac;
            }

            if ((int)per == (int)period - 1)
            {
                double ddbN = Ddb(newTot),
                    slnN = Sl(newTot, per + 1d);
                var isSlnN = straightLine && ddbN < slnN;
                var deprN = isSlnN ? ((int)period == (int)life ? 0d : slnN) : ddbN;
                return newTot + deprN * frac;
            }

            totDepr = newTot;
            per += 1;
        }
    }

    private static double DeprBetween(
        double cost,
        double salvage,
        double life,
        double startPeriod,
        double endPeriod,
        double factor,
        bool straightLine
    ) =>
        TotalDepr(cost, salvage, life, endPeriod, factor, straightLine)
        - TotalDepr(cost, salvage, life, startPeriod, factor, straightLine);

    internal static double Ddb(
        double cost,
        double salvage,
        double life,
        double period,
        double factor
    )
    {
        if ((int)period == 0)
        {
            return Math.Min(cost * (factor / life), cost - salvage);
        }

        return period >= 2d
            ? DeprBetween(cost, salvage, life, period - 1d, period, factor, false)
            : TotalDepr(cost, salvage, life, period, factor, false);
    }

    internal static double Vdb(
        double cost,
        double salvage,
        double life,
        double startPeriod,
        double endPeriod,
        double factor,
        bool noSwitch
    ) => DeprBetween(cost, salvage, life, startPeriod, endPeriod, factor, !noSwitch);

    private static double AmorDaysInYear(DateTime d, int basis) =>
        basis == 1 ? (DateTime.IsLeapYear(d.Year) ? 366d : 365d) : DaysInYear(basis, d, d);

    private static (double firstDepr, double assetLife) FirstDeprLinc(
        double cost,
        DateTime datePurchased,
        DateTime firstPeriod,
        double salvage,
        double rate,
        double assetLife,
        int basis
    )
    {
        DateTime Fix(DateTime d) =>
            (basis == 1 || basis == 3) && DateTime.IsLeapYear(d.Year) && d.Month == 2 && d.Day >= 28
                ? new DateTime(d.Year, d.Month, 28)
                : d;

        var daysInYr = AmorDaysInYear(datePurchased, basis);
        var dp = Fix(datePurchased);
        var fp = Fix(firstPeriod);
        var firstLen = DaysBetween(basis, dp, fp, true);
        var firstDeprTemp = firstLen / daysInYr * rate * cost;
        var firstDepr = firstDeprTemp == 0d ? cost * rate : firstDeprTemp;
        var life = firstDeprTemp == 0d ? assetLife : assetLife + 1d;
        var availDepr = cost - salvage;
        return firstDepr > availDepr ? (availDepr, life) : (firstDepr, life);
    }

    internal static double AmorLinc(
        double cost,
        DateTime datePurchased,
        DateTime firstPeriod,
        double salvage,
        double period,
        double rate,
        int basis
    )
    {
        var assetLifeTemp = Math.Ceiling(1d / rate);

        if (cost == salvage || period > assetLifeTemp)
        {
            return 0d;
        }

        var (firstDepr, _) = FirstDeprLinc(
            cost,
            datePurchased,
            firstPeriod,
            salvage,
            rate,
            assetLifeTemp,
            basis
        );

        if (period == 0d)
        {
            return firstDepr;
        }

        double depr = rate * cost,
            availDepr = cost - salvage - firstDepr,
            counted = 1;

        while (counted <= period)
        {
            depr = depr > availDepr ? availDepr : depr;
            var t = availDepr - depr;
            availDepr = t < 0 ? 0 : t;
            counted += 1;
        }

        return depr;
    }

    private static double DeprCoeff(double assetLife)
    {
        if (assetLife >= 3 && assetLife <= 4)
        {
            return 1.5;
        }

        if (assetLife >= 5 && assetLife <= 6)
        {
            return 2d;
        }

        return assetLife > 6 ? 2.5 : 1d;
    }

    internal static double AmorDegrc(
        double cost,
        DateTime datePurchased,
        DateTime firstPeriod,
        double salvage,
        double period,
        double rate,
        int basis,
        bool excelCompliant
    )
    {
        var assetLife = Math.Ceiling(1d / rate);

        if (cost == salvage || period > assetLife)
        {
            return 0d;
        }

        var deprR = rate * DeprCoeff(assetLife);
        var (fdl, life) = FirstDeprLinc(
            cost,
            datePurchased,
            firstPeriod,
            salvage,
            deprR,
            assetLife,
            basis
        );
        var firstDepr = ExcelRound(fdl, excelCompliant);

        if (period == 0d)
        {
            return firstDepr;
        }

        double depr = 0,
            deprRate = deprR,
            remainCost = cost - firstDepr,
            counted = 1;

        while (true)
        {
            if (counted > period)
            {
                return ExcelRound(depr, excelCompliant);
            }

            counted += 1;
            var calcT = life - counted;
            var atTwo = Math.Abs(calcT - 2d) < 0.0001;
            var deprTemp = atTwo ? remainCost * 0.5 : deprRate * remainCost;

            if (atTwo)
            {
                deprRate = 1d;
            }

            depr =
                remainCost < salvage
                    ? (remainCost - salvage < 0 ? 0 : remainCost - salvage)
                    : deprTemp;
            remainCost -= depr;
        }
    }

    private static double ExcelRound(double x, bool excelCompliant)
    {
        if (excelCompliant)
        {
            var k = Math.Round(x, 13, MidpointRounding.AwayFromZero);
            return Math.Round(k, MidpointRounding.AwayFromZero);
        }

        return Math.Round(x, MidpointRounding.AwayFromZero);
    }

    // ---- rates / value ----

    internal static double Effect(double nominalRate, double npery)
    {
        var periods = Math.Floor(npery);
        return Math.Pow(nominalRate / periods + 1d, periods) - 1d;
    }

    internal static double Nominal(double effectRate, double npery)
    {
        var periods = Math.Floor(npery);
        return (Math.Pow(effectRate + 1d, 1d / periods) - 1d) * periods;
    }

    internal static double Rri(double nper, double pv, double fv) =>
        fv == pv ? 0d : Math.Pow(fv / pv, 1d / nper) - 1d;

    internal static double PDuration(double rate, double pv, double fv) =>
        (Math.Log(fv) - Math.Log(pv)) / Math.Log(1d + rate);

    internal static double ISPmt(double rate, double per, double nper, double pv)
    {
        var coupon = -pv * rate;
        return coupon - coupon / nper * per;
    }

    internal static double FvSchedule(double principal, IReadOnlyList<double> schedule)
    {
        var result = principal;

        foreach (var i in schedule)
        {
            result *= 1d + i;
        }

        return result;
    }

    private static double NpvPeriodic(double rate, IReadOnlyList<double> cashFlows)
    {
        double sum = 0;

        for (var i = 0; i < cashFlows.Count; i++)
        {
            sum += cashFlows[i] / Math.Pow(1d + rate, i + 1);
        }

        return sum;
    }

    internal static double Mirr(
        IReadOnlyList<double> cashFlows,
        double financeRate,
        double reinvestRate
    )
    {
        double n = cashFlows.Count;
        var positives = new List<double>(cashFlows.Count);
        var negatives = new List<double>(cashFlows.Count);

        foreach (var cf in cashFlows)
        {
            positives.Add(cf > 0 ? cf : 0);
            negatives.Add(cf < 0 ? cf : 0);
        }

        return Math.Pow(
                -NpvPeriodic(reinvestRate, positives)
                    * Math.Pow(1d + reinvestRate, n)
                    / (NpvPeriodic(financeRate, negatives) * (1d + financeRate)),
                1d / (n - 1d)
            ) - 1d;
    }

    internal static double XNpv(
        double rate,
        IReadOnlyList<double> cashFlows,
        IReadOnlyList<DateTime> dates
    ) => XNpv(rate, cashFlows, XNpvYearFractions(dates));

    // Per-flow (Days(dates[i], d0) / 365d) exponents, independent of `rate` — computed once and reused
    // across every XIRR solver iteration instead of being recomputed (Days() included) each time.
    private static double[] XNpvYearFractions(IReadOnlyList<DateTime> dates)
    {
        var d0 = dates[0];
        var yearFractions = new double[dates.Count];

        for (var i = 0; i < dates.Count; i++)
        {
            yearFractions[i] = (double)Days(dates[i], d0) / 365d;
        }

        return yearFractions;
    }

    private static double XNpv(double rate, IReadOnlyList<double> cashFlows, double[] yearFractions)
    {
        double sum = 0;

        for (var i = 0; i < cashFlows.Count; i++)
        {
            sum += cashFlows[i] / Math.Pow(1d + rate, yearFractions[i]);
        }

        return sum;
    }

    internal static double XIrr(
        IReadOnlyList<double> cashFlows,
        IReadOnlyList<DateTime> dates,
        double guess
    )
    {
        var yearFractions = XNpvYearFractions(dates);
        return TimeValueOfMoney.Solve(rate => XNpv(rate, cashFlows, yearFractions), guess);
    }

    private static void Dollar(
        double value,
        double fraction,
        out double aBase,
        out double dollar,
        out double remainder,
        out double digits
    )
    {
        aBase = Math.Floor(fraction);
        dollar = value > 0 ? Math.Floor(value) : Math.Ceiling(value);
        remainder = value - dollar;
        digits = Math.Pow(10d, Math.Ceiling(Math.Log10(aBase)));
    }

    internal static double DollarDe(double fractionalDollar, double fraction)
    {
        Dollar(fractionalDollar, fraction, out var b, out var d, out var r, out var g);
        return r * g / b + d;
    }

    internal static double DollarFr(double decimalDollar, double fraction)
    {
        Dollar(decimalDollar, fraction, out var b, out var d, out var r, out var g);
        return r * b / Math.Abs(g) + d;
    }

    // ---- loan cumulative helpers ----
    // The exact annuity forms Excel uses per period, ported so CUMIPMT/CUMPRINC accumulate the same
    // rounding path as the oracle (validated to the last digit by fuzzing).

    private static double LoanFvFactor(double rate, double nper) => Math.Pow(1d + rate, nper);

    private static double LoanPvFactor(double rate, double nper) => 1d / LoanFvFactor(rate, nper);

    private static double LoanAnnuityPv(double rate, double nper, int type) =>
        rate == 0d ? nper : (1d + rate * type) * (1d - LoanPvFactor(rate, nper)) / rate;

    private static double LoanPmt(double rate, double nper, double pv, int type) =>
        -pv / LoanAnnuityPv(rate, nper, type);

    private static double LoanIPmt(double rate, double per, double nper, double pv, int type)
    {
        var result = -(
            pv * LoanFvFactor(rate, per - 1d) * rate
            + LoanPmt(rate, nper, pv, 0) * (LoanFvFactor(rate, per - 1d) - 1d)
        );
        return type == 0 ? result : result / (1d + rate);
    }

    private static double CalcIPmt(double rate, double per, double nper, double pv, int type) =>
        Math.Abs(per - 1d) < 1e-10 && type == 1 ? 0d : LoanIPmt(rate, per, nper, pv, type);

    private static double CalcPPmt(double rate, double per, double nper, double pv, int type) =>
        Math.Abs(per - 1d) < 1e-10 && type == 1
            ? LoanPmt(rate, nper, pv, type)
            : LoanPmt(rate, nper, pv, type) - LoanIPmt(rate, per, nper, pv, type);

    internal static double CumIPmt(
        double rate,
        double nper,
        double pv,
        double startPeriod,
        double endPeriod,
        int type
    )
    {
        double a = 0;

        for (var per = (int)Math.Ceiling(startPeriod); per <= (int)endPeriod; per++)
        {
            a += CalcIPmt(rate, per, nper, pv, type);
        }

        return a;
    }

    internal static double CumPrinc(
        double rate,
        double nper,
        double pv,
        double startPeriod,
        double endPeriod,
        int type
    )
    {
        double a = 0;

        for (var per = (int)Math.Ceiling(startPeriod); per <= (int)endPeriod; per++)
        {
            a += CalcPPmt(rate, per, nper, pv, type);
        }

        return a;
    }

    // ---- odd-period bonds ----

    private static double DaysNotNeg(int basis, DateTime s, DateTime e)
    {
        var r = DaysBetween(basis, s, e, true);
        return r > 0 ? r : 0;
    }

    private static double DaysNotNegHack(int basis, DateTime s, DateTime e)
    {
        if (basis == 0)
        {
            double r = Diff360Us(s, e, Method360.ModifyBothDates);
            return r > 0 ? r : 0;
        }

        return DaysNotNeg(basis, s, e);
    }

    private static double CoupNumberOdd(
        DateTime mat,
        DateTime settl,
        int numMonths,
        bool isWholeNumber
    )
    {
        var couponsTemp = isWholeNumber ? 0d : 1d;
        var endOfMonthTemp = IsLastDayOfMonth(mat);
        var endOfMonth =
            !endOfMonthTemp
            && mat.Month != 2
            && mat.Day > 28
            && mat.Day < DateTime.DaysInMonth(mat.Year, mat.Month)
                ? IsLastDayOfMonth(settl)
                : endOfMonthTemp;
        var startDate = ChangeMonth(settl, 0, endOfMonth);
        var coupons = settl < startDate ? couponsTemp + 1d : couponsTemp;
        var date = ChangeMonth(startDate, numMonths, endOfMonth);
        var (_, _, result) = DatesAggregate1(
            date,
            mat,
            numMonths,
            endOfMonth,
            (_, _) => 1d,
            coupons
        );
        return result;
    }

    /// <summary>
    /// The undocumented-but-required ODDFPRICE/ODDFYIELD precondition: the odd first coupon must line up with
    /// the regular schedule stepping back from maturity (same month/day, February/leap allowances aside).
    /// </summary>
    internal static bool OddFirstCouponAligned(DateTime maturity, DateTime firstCoupon, double freq)
    {
        var numMonthsNeg = -Freq2Months(freq);
        var endMonth = IsLastDayOfMonth(maturity);
        var start = ChangeMonth(maturity, numMonthsNeg, endMonth);
        var (pcd, _) = FindPcdNcd(start, firstCoupon, numMonthsNeg, endMonth);
        return pcd == firstCoupon;
    }

    /// <summary>
    /// Every <see cref="OddFPrice"/> term that doesn't involve <c>yld</c> — the branch choice (short vs.
    /// long first coupon), the coupon-count/day-count walks, and the per-coupon aggregation loop.
    /// Computed once and reused across every ODDFYIELD solver iteration, which otherwise re-ran the whole
    /// walk (including the O(coupons) aggregation loop) just to evaluate a different yield.
    /// </summary>
    private readonly record struct OddFPriceContext(
        bool ShortFirstCoupon,
        double M,
        double Rate,
        double Redemption,
        // short-first-coupon branch
        double N,
        double Y,
        double Term2Const,
        double Term4,
        // long-first-coupon branch
        double Nq,
        double N2,
        double YL,
        double Term2ConstLong,
        double Term4Long
    );

    private static OddFPriceContext BuildOddFPriceContext(
        DateTime settl,
        DateTime mat,
        DateTime issue,
        DateTime firstCoupon,
        double rate,
        double redemption,
        double freq,
        int basis
    )
    {
        var numMonths = Freq2Months(freq);
        var numMonthsNeg = -numMonths;
        var e = CoupDays(basis, settl, firstCoupon, freq);
        var n = CoupNum(settl, mat, freq);
        var m = freq;
        var dfc = DaysNotNeg(basis, issue, firstCoupon);

        if (dfc < e)
        {
            var dsc = DaysNotNeg(basis, settl, firstCoupon);
            var a = DaysNotNeg(basis, issue, settl);
            var y = dsc / e;
            var term2Const = 100d * rate / m * dfc / e;
            var term4 = a / e * (rate / m) * 100d;
            return new OddFPriceContext(
                ShortFirstCoupon: true,
                M: m,
                Rate: rate,
                Redemption: redemption,
                N: n,
                Y: y,
                Term2Const: term2Const,
                Term4: term4,
                Nq: 0,
                N2: 0,
                YL: 0,
                Term2ConstLong: 0,
                Term4Long: 0
            );
        }

        var nc = CoupNum(issue, firstCoupon, freq);
        var lateCoupon = firstCoupon;
        double dcnl = 0,
            anl = 0;

        for (var index = (int)nc; index >= 1; index--)
        {
            var earlyCoupon = ChangeMonth(lateCoupon, numMonthsNeg, false);
            var nl = basis == 1 ? DaysNotNeg(basis, earlyCoupon, lateCoupon) : e;
            var dci = index > 1 ? nl : DaysNotNeg(basis, issue, lateCoupon);
            var startDate = issue > earlyCoupon ? issue : earlyCoupon;
            var endDate = settl < lateCoupon ? settl : lateCoupon;
            var a = DaysNotNeg(basis, startDate, endDate);
            lateCoupon = earlyCoupon;
            dcnl += dci / nl;
            anl += a / nl;
        }

        double dscLong;

        if (basis is 2 or 3)
        {
            var date = CoupNcd(settl, firstCoupon, freq);
            dscLong = DaysNotNeg(basis, settl, date);
        }
        else
        {
            var date = CoupPcd(settl, firstCoupon, freq);
            var a = DaysBetween(basis, date, settl, true);
            dscLong = e - a;
        }

        var nq = CoupNumberOdd(firstCoupon, settl, numMonths, true);
        var n2 = CoupNum(firstCoupon, mat, freq);
        var yL = dscLong / e;
        var term2ConstLong = 100d * rate / m * dcnl;
        var term4Long = 100d * rate / m * anl;

        return new OddFPriceContext(
            ShortFirstCoupon: false,
            M: m,
            Rate: rate,
            Redemption: redemption,
            N: 0,
            Y: 0,
            Term2Const: 0,
            Term4: 0,
            Nq: nq,
            N2: n2,
            YL: yL,
            Term2ConstLong: term2ConstLong,
            Term4Long: term4Long
        );
    }

    // The only part of OddFPrice that depends on `yld` — evaluated once per solver iteration against a
    // context computed once per ODDFYIELD call (see BuildOddFPriceContext).
    private static double EvaluateOddFPrice(in OddFPriceContext ctx, double yld)
    {
        var p1 = yld / ctx.M + 1d;

        if (ctx.ShortFirstCoupon)
        {
            var term1 = ctx.Redemption / Math.Pow(p1, ctx.N - 1d + ctx.Y);
            var term2 = ctx.Term2Const / Math.Pow(p1, ctx.Y);
            double term3 = 0;

            for (var index = 2; index <= (int)ctx.N; index++)
            {
                term3 += 100d * ctx.Rate / ctx.M / Math.Pow(p1, index - 1d + ctx.Y);
            }

            return term1 + term2 + term3 - ctx.Term4;
        }

        var t1 = ctx.Redemption / Math.Pow(p1, ctx.YL + ctx.Nq + ctx.N2);
        var t2 = ctx.Term2ConstLong / Math.Pow(p1, ctx.Nq + ctx.YL);
        double t3 = 0;

        for (var index = 1; index <= (int)ctx.N2; index++)
        {
            t3 += 100d * ctx.Rate / ctx.M / Math.Pow(p1, index + ctx.Nq + ctx.YL);
        }

        return t1 + t2 + t3 - ctx.Term4Long;
    }

    internal static double OddFPrice(
        DateTime settl,
        DateTime mat,
        DateTime issue,
        DateTime firstCoupon,
        double rate,
        double yld,
        double redemption,
        double freq,
        int basis
    )
    {
        var ctx = BuildOddFPriceContext(
            settl,
            mat,
            issue,
            firstCoupon,
            rate,
            redemption,
            freq,
            basis
        );
        return EvaluateOddFPrice(ctx, yld);
    }

    internal static double OddFYield(
        DateTime settl,
        DateTime mat,
        DateTime issue,
        DateTime firstCoupon,
        double rate,
        double pr,
        double redemption,
        double freq,
        int basis
    )
    {
        var years = DaysBetween(basis, settl, mat, true);
        var px = pr - 100d;
        var num = rate * years * 100d - px;
        var denum = px / 4d + years * px / 2d + years * 100d;
        var guess = num / denum;
        var ctx = BuildOddFPriceContext(
            settl,
            mat,
            issue,
            firstCoupon,
            rate,
            redemption,
            freq,
            basis
        );
        return TimeValueOfMoney.Solve(yld => pr - EvaluateOddFPrice(ctx, yld), guess);
    }

    internal static double OddLFunc(
        DateTime settl,
        DateTime mat,
        DateTime lastInterest,
        double rate,
        double priceOrYield,
        double redemption,
        double freq,
        int basis,
        bool isLPrice
    )
    {
        var m = freq;
        var numMonths = (int)(12d / freq);
        var nc = CoupNum(lastInterest, mat, freq);
        var earlyCoupon = lastInterest;
        double dcnl = 0,
            anl = 0,
            dscnl = 0;

        for (var index = 1; index <= (int)nc; index++)
        {
            var lateCoupon = ChangeMonth(earlyCoupon, numMonths, false);
            var nl = DaysNotNegHack(basis, earlyCoupon, lateCoupon);
            var dci = index < (int)nc ? nl : DaysNotNegHack(basis, earlyCoupon, mat);
            var a =
                lateCoupon < settl ? dci
                : earlyCoupon < settl ? DaysNotNeg(basis, earlyCoupon, settl)
                : 0d;
            var startDate = settl > earlyCoupon ? settl : earlyCoupon;
            var endDate = mat < lateCoupon ? mat : lateCoupon;
            var dsc = DaysNotNeg(basis, startDate, endDate);
            earlyCoupon = lateCoupon;
            dcnl += dci / nl;
            anl += a / nl;
            dscnl += dsc / nl;
        }

        var x = 100d * rate / m;
        var term1 = dcnl * x + redemption;

        if (isLPrice)
        {
            return term1 / (dscnl * priceOrYield / m + 1d) - anl * x;
        }

        var term2 = anl * x + priceOrYield;
        var term3 = m / dscnl;
        return (term1 - term2) / term2 * term3;
    }
}
