using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace Danfma.MySheet.Expressions.Text;

/// <summary>
/// A process-wide cache of compiled <see cref="Regex"/> instances, shared by every wildcard/regex site in the
/// engine (REGEXTEST/REGEXEXTRACT/REGEXREPLACE, SEARCH's wildcard translation, and the SUMIF/COUNTIF/XLOOKUP
/// wildcard criteria). Without it, a column of 100k REGEXTEST cells with the SAME pattern recompiles the regex
/// 100k times — <see cref="Regex"/> construction (parsing + code-gen) dominates the per-cell cost far more than
/// the match itself.
///
/// <para>Eviction is deliberately dumb: once the cache holds <see cref="Capacity"/> distinct (pattern, options)
/// keys, the NEXT miss clears it wholesale and starts over. Real spreadsheets reuse a handful of stable
/// patterns per column/criterion, so this cap is essentially never reached; an LRU or size-aware policy would
/// add locking/bookkeeping for a scenario (thousands of DISTINCT patterns live at once) that does not happen in
/// practice. A full clear under that pathological case just pays a few extra recompiles — the same cost as
/// having no cache at all.</para>
///
/// <para>Thread-safe by construction: <see cref="ConcurrentDictionary{TKey,TValue}"/>'s
/// <c>GetOrAdd</c> guarantees a single published <see cref="Regex"/> instance per key even under concurrent
/// evaluation (parallel column scans). The eviction check races benignly — at worst two threads both decide to
/// clear, which is still correct and still cheap.</para>
/// </summary>
internal static class RegexCache
{
    // Real workbooks reuse a handful of stable patterns per column/criterion, so this cap is essentially never
    // hit on real data; see the class doc for why a wholesale clear (rather than LRU) is fine when it is.
    private const int Capacity = 256;

    private static readonly ConcurrentDictionary<
        (string Pattern, RegexOptions Options),
        Regex
    > Cache = new();

    /// <summary>
    /// Returns the compiled <see cref="Regex"/> for (<paramref name="pattern"/>, <paramref name="options"/>),
    /// building and caching it on first use. Every entry is built with a defensive 1-second match timeout
    /// (<see cref="TimeSpan.FromSeconds(double)"/>) — protection against ReDoS/catastrophic backtracking on
    /// patterns that ultimately come from a formula — regardless of whether it was already cached. Throws
    /// <see cref="ArgumentException"/> for an invalid pattern, exactly like <c>new Regex(...)</c> would.
    /// </summary>
    public static Regex Get(string pattern, RegexOptions options)
    {
        var key = (pattern, options);

        if (Cache.Count >= Capacity && !Cache.ContainsKey(key))
        {
            Cache.Clear();
        }

        return Cache.GetOrAdd(
            key,
            static k => new Regex(k.Pattern, k.Options, TimeSpan.FromSeconds(1))
        );
    }
}
