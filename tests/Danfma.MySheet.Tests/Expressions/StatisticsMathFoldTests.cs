using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Direct contract of the four population folds in <see cref="StatisticsMath"/> — Median, Mode,
/// QuartileInclusive and QuartileExclusive — extracted from MEDIAN/MODE.SNGL/QUARTILE.INC/QUARTILE.EXC.
/// AGGREGATE's codes 12/13/17/19 reuse them over an already collected population, so the contract is
/// pinned here on plain lists, independent of any node's argument handling. The end-to-end behaviour of
/// the four functions stays pinned by OrderStatisticTests.
/// </summary>
public class StatisticsMathFoldTests
{
    // Turns "the fold was expected to fail" into a readable failure instead of a null-deref.
    private static Error Failure(Error? error) =>
        error ?? throw new InvalidOperationException("expected the fold to fail, but it succeeded");

    // --- Median (MEDIAN) — takes an ASCENDING sorted population ---

    [Test]
    public async Task Median_Empty_IsNumError()
    {
        await Assert.That(Failure(StatisticsMath.Median([], out _))).IsEqualTo(Error.Num);
    }

    [Test]
    public async Task Median_OddCount_IsTheMiddleValue()
    {
        // 5 values -> index 2. Deliberately skewed so the middle value (3) is NOT the mean (4.4).
        var error = StatisticsMath.Median([1.0, 2.0, 3.0, 4.0, 12.0], out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(3.0);
    }

    [Test]
    public async Task Median_EvenCount_AveragesTheTwoMiddleValues()
    {
        // 4 values -> mean of indexes 1 and 2 = (2+5)/2 = 3.5, not the mean of all four (27).
        var error = StatisticsMath.Median([1.0, 2.0, 5.0, 100.0], out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(3.5);
    }

    // --- Mode (MODE.SNGL) — takes the population in SCAN ORDER, never sorted ---

    [Test]
    public async Task Mode_NoRepeatedValue_IsNAError()
    {
        // Documented rule: "If the data set contains no duplicate data points, MODE.SNGL returns #N/A."
        await Assert
            .That(Failure(StatisticsMath.Mode([1.0, 2.0, 3.0, 4.0], out _)))
            .IsEqualTo(Error.NA);
    }

    [Test]
    public async Task Mode_MostFrequentValue_Wins()
    {
        var error = StatisticsMath.Mode([5.6, 4.0, 4.0, 3.0, 2.0, 4.0], out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(4.0);
    }

    [Test]
    public async Task Mode_Tie_TakesTheFirstValueToReachTheWinningCount()
    {
        // [2,1,1,2]: both values end at count 2, but 1 reaches 2 first (index 2) and the strict '>'
        // never lets a later value take the crown. So the tie-break is "first to REACH the winning
        // count", NOT "first value encountered" — which would answer 2 here.
        var error = StatisticsMath.Mode([2.0, 1.0, 1.0, 2.0], out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(1.0);
    }

    [Test]
    public async Task Mode_DoesNotSortThePopulation()
    {
        // [3,5,5,3]: in scan order 5 reaches count 2 first -> 5. Sorting first ([3,3,5,5]) would answer
        // 3, so this pins that the fold consumes the population exactly as it was scanned.
        var error = StatisticsMath.Mode([3.0, 5.0, 5.0, 3.0], out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(5.0);
    }

    // --- QuartileInclusive (QUARTILE.INC) — quart 0-4 over an ascending sorted population ---

    [Test]
    public async Task QuartileInclusive_QuartBelowZero_IsNumError()
    {
        await Assert
            .That(Failure(StatisticsMath.QuartileInclusive([1.0, 2.0, 3.0, 4.0], -1, out _)))
            .IsEqualTo(Error.Num);
    }

    [Test]
    public async Task QuartileInclusive_QuartAboveFour_IsNumError()
    {
        await Assert
            .That(Failure(StatisticsMath.QuartileInclusive([1.0, 2.0, 3.0, 4.0], 5, out _)))
            .IsEqualTo(Error.Num);
    }

    [Test]
    public async Task QuartileInclusive_QuartTwo_IsTheMedian()
    {
        // quart 2 -> PERCENTILE.INC(k = 0.5) -> the median of [1,2,3,4] = 2.5.
        var error = StatisticsMath.QuartileInclusive([1.0, 2.0, 3.0, 4.0], 2, out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(2.5);
    }

    [Test]
    public async Task QuartileInclusive_TruncatesQuart()
    {
        // 2.9 truncates to 2, so it must answer the median, not interpolate at 2.9/4.
        var error = StatisticsMath.QuartileInclusive([1.0, 2.0, 3.0, 4.0], 2.9, out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(2.5);
    }

    [Test]
    public async Task QuartileInclusive_EmptyPopulation_IsNumError()
    {
        await Assert
            .That(Failure(StatisticsMath.QuartileInclusive([], 2, out _)))
            .IsEqualTo(Error.Num);
    }

    // --- QuartileExclusive (QUARTILE.EXC) — NO range check of its own ---

    [Test]
    public async Task QuartileExclusive_QuartZero_IsNumError()
    {
        // 0/4 = 0 is rejected by PercentileExclusive's own "k is <= 0 or >= 1" guard, not by a range
        // check in the fold: the exclusive definition cannot reach the 0th percentile.
        await Assert
            .That(Failure(StatisticsMath.QuartileExclusive([1.0, 5.0, 9.0], 0, out _)))
            .IsEqualTo(Error.Num);
    }

    [Test]
    public async Task QuartileExclusive_QuartFour_IsNumError()
    {
        await Assert
            .That(Failure(StatisticsMath.QuartileExclusive([1.0, 5.0, 9.0], 4, out _)))
            .IsEqualTo(Error.Num);
    }

    [Test]
    public async Task QuartileExclusive_QuartTwo_IsTheMedian()
    {
        // quart 2 -> PERCENTILE.EXC(k = 0.5) -> rank 0.5*(3+1) = 2 -> sorted[1] = 5.
        var error = StatisticsMath.QuartileExclusive([1.0, 5.0, 9.0], 2, out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(5.0);
    }

    [Test]
    public async Task QuartileExclusive_TruncatesQuart()
    {
        var error = StatisticsMath.QuartileExclusive([1.0, 5.0, 9.0], 2.9, out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(5.0);
    }

    [Test]
    public async Task QuartileExclusive_QuartTwoOverASingleValue_IsThatValue()
    {
        // Measured, not assumed: n = 1 gives rank 0.5*(1+1) = 1, which is exactly reachable, so the
        // single value comes back — the exclusive form is NOT #NUM! here.
        var error = StatisticsMath.QuartileExclusive([7.0], 2, out var result);

        await Assert.That(error.HasValue).IsFalse();
        await Assert.That(result).IsEqualTo(7.0);
    }

    [Test]
    public async Task QuartileExclusive_UnreachableRank_IsNumError()
    {
        // n = 2, quart 1 -> rank 0.25*(2+1) = 0.75 < 1: the first quartile lies below the smallest
        // data point, which the exclusive definition reports as #NUM!.
        await Assert
            .That(Failure(StatisticsMath.QuartileExclusive([1.0, 5.0], 1, out _)))
            .IsEqualTo(Error.Num);
    }
}
