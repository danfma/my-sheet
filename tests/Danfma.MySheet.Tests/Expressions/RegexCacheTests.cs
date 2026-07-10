using System.Text.RegularExpressions;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Expressions.Text;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The shared <see cref="RegexCache"/> behind REGEXTEST/REGEXEXTRACT/REGEXREPLACE, SEARCH's wildcard
/// translation and the SUMIF/COUNTIF/XLOOKUP wildcard criteria: same (pattern, options) must return the
/// SAME compiled instance, the cache must survive well past its documented capacity without breaking, and
/// the 1-second defensive timeout must still fire on a catastrophic pattern.
/// </summary>
public class RegexCacheTests
{
    [Test]
    public async Task Get_SamePatternAndOptions_ReturnsSameInstance()
    {
        var first = RegexCache.Get(@"^[a-z]+\d*$", RegexOptions.CultureInvariant);
        var second = RegexCache.Get(@"^[a-z]+\d*$", RegexOptions.CultureInvariant);

        await Assert.That(ReferenceEquals(first, second)).IsTrue();
    }

    [Test]
    public async Task Get_SamePatternDifferentOptions_ReturnsDifferentInstances()
    {
        var caseSensitive = RegexCache.Get("abc", RegexOptions.None);
        var caseInsensitive = RegexCache.Get("abc", RegexOptions.IgnoreCase);

        await Assert.That(ReferenceEquals(caseSensitive, caseInsensitive)).IsFalse();
    }

    [Test]
    public async Task Get_BeyondCapacity_DoesNotExplodeAndStaysFunctional()
    {
        // 300 distinct patterns > the cache's documented 256-entry cap: this forces at least one wholesale
        // eviction (see RegexCache's doc for why that is the deliberate, cheap policy). The cache must keep
        // working correctly through and after the eviction.
        for (var i = 0; i < 300; i++)
        {
            var regex = RegexCache.Get($"^pattern{i}$", RegexOptions.None);

            await Assert.That(regex.IsMatch($"pattern{i}")).IsTrue();
            await Assert.That(regex.IsMatch($"pattern{i}x")).IsFalse();
        }

        // The cache is usable afterwards, including re-resolving a pattern evicted by the sweep.
        var again = RegexCache.Get("^pattern0$", RegexOptions.None);
        await Assert.That(again.IsMatch("pattern0")).IsTrue();
    }

    [Test]
    public async Task Get_CatastrophicPattern_StillTimesOut()
    {
        // Classic catastrophic-backtracking shape ((a+)+) against a long non-matching input: without the
        // 1s timeout this would hang. The cache must preserve RegexFunctions'/Criteria's defensive timeout
        // regardless of caching.
        var regex = RegexCache.Get("(a+)+$", RegexOptions.None);
        var input = new string('a', 40) + "!";

        await Assert.That(() => regex.Match(input)).Throws<RegexMatchTimeoutException>();
    }

    [Test]
    public async Task WildcardMatch_CatastrophicPattern_FailsSafeInsteadOfThrowing()
    {
        // Criteria.WildcardMatch (the static path XLOOKUP/XMATCH/LOOKUP wildcard mode uses, see
        // LookupMatching) now routes through RegexCache too, which means it inherits the 1s timeout the
        // old un-timed `Regex.IsMatch` call never had. A pathological run of '*' wildcards translates to
        // an adjacent ".*" chain — the same catastrophic-backtracking shape as above — so WildcardMatch
        // must catch the timeout and fail safe as "no match" rather than let the exception escape into a
        // *IFS/XLOOKUP scan that never expected one.
        var pattern = string.Concat(Enumerable.Repeat("a*", 25)) + "b";
        var text = new string('a', 40);

        var matched = Criteria.WildcardMatch(pattern, text);

        await Assert.That(matched).IsFalse();
    }
}
