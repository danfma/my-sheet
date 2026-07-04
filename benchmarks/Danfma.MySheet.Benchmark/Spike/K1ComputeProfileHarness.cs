using System.Collections.Concurrent;
using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MySheetWorkbook = Danfma.MySheet.Workbook;

namespace Danfma.MySheet.Benchmark.Spike;

// Compute-phase profiler for the K1 shape, to localize where the ~687ms of MySheet's compute pass live and
// to put a MEASURED upper bound on the "numeric addressing" (4.0) redesign. INVESTIGATION ONLY — nothing in
// Danfma.MySheet is touched; every number here is produced by benchmark-side probes that either (a) drive the
// real engine, or (b) micro-benchmark an isolated component with the SAME call counts the engine incurs.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --k1-compute-attrib
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --k1-compute-loop [N]   (for profiler)
//
// SHAPE (identical to K1EndToEndHarness):
//   Data: 100k rows x 2 cols. A = "Show"/"Hide" (~1/7 Show); B = numbers (B{v}=v).
//   S: 400k formulas from 2 skeletons, cross-sheet, v=(r%100k)+1:
//        C{r} = IF(Data!A{v}="Show", Data!B{v}*2, Data!B{v}/2)
//        D{r} = Data!B{v}*2+1
//
// CALL MODEL (derived from the AST, cross-checked by ProbeCallCount against the live engine's cache size):
//   GetCellValue calls in one compute sweep:
//     - 400k  self reads (the sweep loop over C{r} and D{r})
//     - 400k  from C: condition ref Data!A{v} + taken-branch ref Data!B{v} (2 per C, 200k C cells)
//     - 200k  from D: Data!B{v} (1 per D, 200k D cells)
//     = 1.00M GetCellValue calls total.
//   Distinct cells (cache entries after a sweep): 400k S + A1..A100k + B1..B100k = 600k -> 600k are MISSES,
//   400k are cache HITS (repeat Data refs, since v cycles).
//   Per GetCellValue the (string,string) tuple key is hashed:
//     - every call: 1x cache.TryGetValue
//     - every miss (600k): evaluating.Add + cache store + evaluating.Remove
//     = 1.00M + 3*600k = 2.80M tuple hash/equals ops.
//   Plus per miss (600k): 1x Sheets.TryGetValue(name) [OrdinalIgnoreCase string hash] and 1x _cells
//   TryGetValue(id) [ordinal string hash].
public static class K1ComputeProfileHarness
{
    private const string DataSheet = "Data";
    private const string FormulaSheet = "S";
    private const int DataRows = 100_000;
    private const int FormulaRows = 200_000; // C{r} and D{r} -> 400k formulas
    private const int ShowEvery = 7;

    private static Expression Show { get; } = new StringValue("Show");
    private static Expression Hide { get; } = new StringValue("Hide");

    // ============================================================ compute-only loop (for external profiler)

    // Build once, then InvalidateCache + full sweep N times so a CPU sampler sees the eval path dominate.
    public static void RunLoop(string[] args)
    {
        var n = 20;
        var idx = Array.IndexOf(args, "--k1-compute-loop");
        if (idx >= 0 && idx + 1 < args.Length && int.TryParse(args[idx + 1], out var parsed))
        {
            n = parsed;
        }

        Console.WriteLine(
            $"== K1 compute-only loop: build once, then {n}x (InvalidateCache + sweep) =="
        );
        var wb = BuildFull();
        var warm = SweepCompute(wb); // warm JIT + first population
        Console.WriteLine($"warm sweep: agg {warm.Aggregate:R}, {warm.Ms:N1} ms");

        var best = double.MaxValue;
        double lastAgg = 0;
        for (var i = 0; i < n; i++)
        {
            wb.InvalidateCache();
            var r = SweepCompute(wb);
            lastAgg = r.Aggregate;
            best = Math.Min(best, r.Ms);
            Console.WriteLine($"  iter {i, 2}: {r.Ms, 7:N1} ms");
        }

        Console.WriteLine($"best sweep: {best:N1} ms (agg {lastAgg:R})");
    }

    // ============================================================ allocation attribution (Phase 2)

    // Attributes the ~42.7 MB the K1EndToEndHarness reports for the compute phase, isolating the ENGINE's
    // transient allocation from two things that ride inside that phase's measured region but are NOT the eval
    // path: (1) the harness building the id strings "C"+r / "D"+r (an int.ToString temp + the concat result per
    // call), and (2) the dense value store's own retained page backing (ComputedValue[1024] + a 128-byte
    // presence bitmap per page). Each probe is best-of-N (min) with GcQuiesce between trials.
    //
    //   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --k1-compute-alloc
    public static void RunAllocAttribution()
    {
        Console.WriteLine("== K1 compute allocation attribution — directed probes ==");
        Console.WriteLine(
            $"Runtime {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"Shape: Data {DataRows:N0}x2, S {FormulaRows * 2:N0} formulas. "
                + "Each probe best-of-5 (min); GC.Collect between trials; alloc via GetTotalAllocatedBytes(true).");
        Console.WriteLine();

        // Pre-build the id strings ONCE, outside every measured region, so a probe can sweep the engine without
        // the "C"+r / "D"+r building landing in its allocation reading.
        var cIds = new string[FormulaRows];
        var dIds = new string[FormulaRows];
        for (var r = 1; r <= FormulaRows; r++)
        {
            cIds[r - 1] = "C" + r;
            dIds[r - 1] = "D" + r;
        }

        var wb = BuildFull();
        _ = SweepCompute(wb); // warm JIT + first population

        // Probe A — live compute sweep, harness building "C"+r inline (the SAME shape K1EndToEndHarness measures).
        var allocA = BestAlloc(() =>
        {
            wb.InvalidateCache();
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            _ = SweepCompute(wb);
            return GC.GetTotalAllocatedBytes(true) - before;
        });

        // Probe B — id-string building ONLY, no engine touch: the harness measurement artifact per sweep.
        var allocB = BestAlloc(() =>
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            long sink = 0;
            for (var r = 1; r <= FormulaRows; r++)
            {
                sink += ("C" + r).Length;
                sink += ("D" + r).Length;
            }

            GC.KeepAlive(sink);
            return GC.GetTotalAllocatedBytes(true) - before;
        });

        // Probe C — compute sweep over PRE-BUILT ids (measured region = pure engine: store page backing that a
        // cold sweep must allocate + any eval-path transient).
        var allocC = BestAlloc(() =>
        {
            wb.InvalidateCache();
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            _ = SweepPrebuilt(wb, cIds, dIds);
            return GC.GetTotalAllocatedBytes(true) - before;
        });

        // Probe D — warm re-sweep over PRE-BUILT ids (every cell already memoized: all cache HITS). This is the
        // pure read path; it allocates nothing if the memoized-read path is alloc-free.
        _ = SweepPrebuilt(wb, cIds, dIds); // ensure warm
        var allocD = BestAlloc(() =>
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            _ = SweepPrebuilt(wb, cIds, dIds);
            return GC.GetTotalAllocatedBytes(true) - before;
        });

        double Mb(long b) => b / 1048576d;

        Console.WriteLine("-- raw probe numbers (per one compute sweep) --");
        Console.WriteLine($"A  live sweep, harness builds \"C\"+r inline   : {Mb(allocA), 8:N1} MB");
        Console.WriteLine($"B  id-string building only (no engine)        : {Mb(allocB), 8:N1} MB");
        Console.WriteLine($"C  cold sweep, PRE-BUILT ids (pure engine)    : {Mb(allocC), 8:N1} MB");
        Console.WriteLine($"D  warm re-sweep, PRE-BUILT ids (all hits)    : {Mb(allocD), 8:N1} MB");
        Console.WriteLine();

        // The cold engine sweep (C) that is NOT read-path transient (D) is the store's page backing a first
        // population must allocate — dense arrays that REPLACED the old 1-heap-Node-per-entry dictionary.
        var storePages = allocC - allocD;
        var evalTransient = allocD;
        var harnessStrings = allocB;

        Console.WriteLine(new string('=', 72));
        Console.WriteLine($"{"Compute allocation attribution", -44}{"MB", 10}{"% of A", 14}");
        Console.WriteLine(new string('-', 72));
        AllocRow("harness id-string building (\"C\"+r)", harnessStrings, allocA);
        AllocRow("dense store page backing (cold populate)", storePages, allocA);
        AllocRow("engine read/eval transient (warm sweep)", evalTransient, allocA);
        Console.WriteLine(new string('-', 72));
        AllocRow("= accounted (B + C)", allocB + allocC, allocA);
        AllocRow("live compute sweep (A, endtoend figure)", allocA, allocA);
        Console.WriteLine(new string('=', 72));
        Console.WriteLine();
        Console.WriteLine(
            $"ENGINE compute allocation (excluding the harness's \"C\"+r artifact) = C = {Mb(allocC):N1} MB, "
                + $"of which read/eval transient = {Mb(evalTransient):N1} MB and dense-store page backing = "
                + $"{Mb(storePages):N1} MB (retained, intrinsic to a dense store).");
    }

    private static long BestAlloc(Func<long> probe)
    {
        var best = long.MaxValue;
        for (var t = 0; t < 5; t++)
        {
            best = Math.Min(best, probe());
        }

        return best;
    }

    private static void AllocRow(string label, long bytes, long total)
    {
        var mb = bytes / 1048576d;
        Console.WriteLine($"{label, -44}{mb, 10:N1}{(double)bytes / total * 100, 13:N1}%");
    }

    private static SweepResult SweepPrebuilt(MySheetWorkbook wb, string[] cIds, string[] dIds)
    {
        var sw = Stopwatch.StartNew();
        double sum = 0;
        for (var i = 0; i < cIds.Length; i++)
        {
            if (wb.GetCellValue(FormulaSheet, cIds[i]).TryGetNumber(out var c))
            {
                sum += c;
            }

            if (wb.GetCellValue(FormulaSheet, dIds[i]).TryGetNumber(out var d))
            {
                sum += d;
            }
        }

        sw.Stop();
        return new SweepResult(sw.Elapsed.TotalMilliseconds, sum);
    }

    // ============================================================ attribution

    public static void RunAttribution()
    {
        Console.WriteLine("== K1 compute attribution — directed probes ==");
        Console.WriteLine(
            $"Runtime {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"Shape: Data {DataRows:N0}x2, S {FormulaRows * 2:N0} formulas (C=IF cross-sheet, D=arith). "
                + "Each probe best-of-N (min); GC.Collect between trials."
        );
        Console.WriteLine();

        var full = ProbeFullCompute();
        var callModel = ProbeCallCount(full.Workbook);
        var keys = ProbeKeyHashing();
        var sheetRes = ProbeSheetResolution();
        var split = ProbeSkeletonSplit();
        var evalOnly = ProbeAstOnly();

        // ---- raw probe dump ----
        Console.WriteLine("-- raw probe numbers --");
        Console.WriteLine(
            $"full compute sweep (live engine)     : {full.Ms, 8:N1} ms  ({full.Alloc / 1048576d:N1} MB)"
        );
        Console.WriteLine(
            $"  GetCellValue calls (model)         : {callModel.TotalCalls, 12:N0}  (miss {callModel.Misses:N0} / hit {callModel.Hits:N0})"
        );
        Console.WriteLine(
            $"  cache size after sweep (live)      : {callModel.LiveCacheSize, 12:N0}  (model distinct {callModel.Misses:N0})"
        );
        Console.WriteLine($"  tuple key ops (model)              : {callModel.TupleOps, 12:N0}");
        Console.WriteLine();
        Console.WriteLine(
            $"key hashing replay  (string,string)  : {keys.StringMs, 8:N1} ms   over {keys.Ops:N0} ops"
        );
        Console.WriteLine(
            $"key hashing replay  (int,int)        : {keys.IntMs, 8:N1} ms   over {keys.Ops:N0} ops"
        );
        Console.WriteLine(
            $"  -> string-key overhead vs int-key  : {keys.StringMs - keys.IntMs, 8:N1} ms"
        );
        Console.WriteLine();
        Console.WriteLine($"sheet-by-name resolution  (600k)     : {sheetRes.ByNameMs, 8:N1} ms");
        Console.WriteLine($"sheet handle deref        (600k)     : {sheetRes.ByHandleMs, 8:N1} ms");
        Console.WriteLine(
            $"  -> sheet-name overhead             : {sheetRes.ByNameMs - sheetRes.ByHandleMs, 8:N1} ms"
        );
        Console.WriteLine();
        Console.WriteLine(
            $"skeleton split — C only (IF)         : {split.CMs, 8:N1} ms   ({split.CMs / (FormulaRows) * 1000:N3} us/formula)"
        );
        Console.WriteLine(
            $"skeleton split — D only (arith)      : {split.DMs, 8:N1} ms   ({split.DMs / (FormulaRows) * 1000:N3} us/formula)"
        );
        Console.WriteLine();
        Console.WriteLine(
            $"AST-only eval (no GetCellValue plumb): {evalOnly.Ms, 8:N1} ms   over {evalOnly.Calls:N0} Evaluate() calls"
        );
        Console.WriteLine();

        // ---- attribution table ----
        // The compute sweep = key hashing + sheet resolution + _cells id lookup + AST-node walk + coercion + GC.
        // The (int,int)+handle redesign removes exactly [string-key overhead] + [sheet-name overhead] + [id
        // string-lookup overhead]. We measure the first two directly; the third is folded into the AST/plumbing
        // remainder and separated by the id-lookup probe below.
        var idLookup = ProbeIdLookup();
        Console.WriteLine($"id string-lookup in _cells (600k)    : {idLookup.StringMs, 8:N1} ms");
        Console.WriteLine($"id numeric-index deref     (600k)    : {idLookup.IntMs, 8:N1} ms");
        Console.WriteLine(
            $"  -> id string-lookup overhead       : {idLookup.StringMs - idLookup.IntMs, 8:N1} ms"
        );
        Console.WriteLine();

        // Structural headroom: the SAME cache traffic (1.0M TryGetValue + 600k insert, no evaluating churn) on
        // four backing structures. This separates the STRING-KEY tax (captured above) from the CONCURRENT-DICT
        // STRUCTURE tax that int keys alone do NOT remove — only a denser store (plain Dictionary / flat array,
        // reachable once cells are numerically addressed) does.
        var cacheS = ProbeCacheStructure();
        Console.WriteLine("-- cache-structure headroom (1.0M lookup + 600k insert, pure cache) --");
        Console.WriteLine(
            $"  ConcurrentDictionary<(string,string)> : {cacheS.CDictStr, 8:N1} ms   (today)"
        );
        Console.WriteLine($"  ConcurrentDictionary<(int,int)>       : {cacheS.CDictInt, 8:N1} ms");
        Console.WriteLine($"  Dictionary<(int,int)>                 : {cacheS.DictInt, 8:N1} ms");
        Console.WriteLine($"  flat array[] (dense numeric address)  : {cacheS.FlatArray, 8:N1} ms");
        Console.WriteLine();
        var churn = ProbeEvaluatingChurn();
        Console.WriteLine(
            $"evaluating-set churn (string) 600k Add+Remove : {churn.StringMs, 8:N1} ms"
        );
        Console.WriteLine(
            $"evaluating-set churn (int)    600k Add+Remove : {churn.IntMs, 8:N1} ms"
        );
        Console.WriteLine($"  (a visited-bit on the cell would remove this entirely)");
        Console.WriteLine();

        var total = full.Ms;
        var keyGain = keys.StringMs - keys.IntMs;
        var sheetGain = sheetRes.ByNameMs - sheetRes.ByHandleMs;
        var idGain = idLookup.StringMs - idLookup.IntMs;
        var numericGain = keyGain + sheetGain + idGain;
        var remainder = total - numericGain;

        Console.WriteLine(new string('=', 72));
        Console.WriteLine($"{"Compute cost attribution", -40}{"ms", 10}{"% of compute", 14}");
        Console.WriteLine(new string('-', 72));
        AttribRow("string-tuple key hash/equals", keyGain, total);
        AttribRow("sheet resolution by name", sheetGain, total);
        AttribRow("cell-id string lookup in _cells", idGain, total);
        AttribRow("  = numeric-addressing removable", numericGain, total);
        AttribRow("AST walk + coercion + GC (remainder)", remainder, total);
        Console.WriteLine(new string('-', 72));
        AttribRow("TOTAL compute (live)", total, total);
        Console.WriteLine(new string('=', 72));
        Console.WriteLine();
        Console.WriteLine(
            $"Measured UPPER BOUND of numeric addressing over the {total:N0}ms compute: "
                + $"~{numericGain:N0} ms ({numericGain / total * 100:N0}%)."
        );
    }

    private static void AttribRow(string label, double ms, double total)
    {
        Console.WriteLine($"{label, -40}{ms, 10:N1}{ms / total * 100, 13:N1}%");
    }

    // ------------------------------------------------------------ live full compute

    private readonly record struct FullResult(double Ms, long Alloc, MySheetWorkbook Workbook);

    private static FullResult ProbeFullCompute()
    {
        var wb = BuildFull();
        _ = SweepCompute(wb); // warm + populate
        var best = double.MaxValue;
        long alloc = 0;
        for (var i = 0; i < 3; i++)
        {
            wb.InvalidateCache();
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var r = SweepCompute(wb);
            alloc = GC.GetTotalAllocatedBytes(true) - before;
            best = Math.Min(best, r.Ms);
        }

        // Leave the cache warm so ProbeCallCount can read the live cache size.
        _ = SweepCompute(wb);
        return new FullResult(best, alloc, wb);
    }

    private readonly record struct SweepResult(double Ms, double Aggregate);

    private static SweepResult SweepCompute(MySheetWorkbook wb)
    {
        var sw = Stopwatch.StartNew();
        double sum = 0;
        for (var r = 1; r <= FormulaRows; r++)
        {
            if (wb.GetCellValue(FormulaSheet, "C" + r).TryGetNumber(out var c))
            {
                sum += c;
            }

            if (wb.GetCellValue(FormulaSheet, "D" + r).TryGetNumber(out var d))
            {
                sum += d;
            }
        }

        sw.Stop();
        return new SweepResult(sw.Elapsed.TotalMilliseconds, sum);
    }

    // ------------------------------------------------------------ call-count model + live cross-check

    private readonly record struct CallModel(
        long TotalCalls,
        long Misses,
        long Hits,
        long TupleOps,
        long LiveCacheSize
    );

    private static CallModel ProbeCallCount(MySheetWorkbook wb)
    {
        // Derived from the AST (see class header).
        long selfCalls = FormulaRows * 2; // 400k
        long refCallsFromC = FormulaRows * 2L; // Data!A + taken Data!B, per C  -> 400k
        long refCallsFromD = FormulaRows * 1L; // Data!B per D                  -> 200k
        long total = selfCalls + refCallsFromC + refCallsFromD; // 1.00M
        long misses = (long)FormulaRows * 2 + DataRows * 2L; // 400k S + 200k Data distinct = 600k
        long hits = total - misses; // 400k
        long tupleOps = total + 3 * misses; // TryGetValue + (Add+store+Remove) per miss

        // Cross-check: the live cache size after a warm sweep must equal the distinct-cell count (misses).
        // Read the private _cache via reflection (benchmark-side only; no production surface added).
        long liveCacheSize = ReadCacheCount(wb);

        return new CallModel(total, misses, hits, tupleOps, liveCacheSize);
    }

    private static long ReadCacheCount(MySheetWorkbook wb)
    {
        var field = typeof(MySheetWorkbook).GetField(
            "_cache",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance
        );
        var cache = field?.GetValue(wb);
        if (cache is System.Collections.ICollection c)
        {
            return c.Count;
        }

        return -1;
    }

    // ------------------------------------------------------------ key hashing replay (string vs int)

    private readonly record struct KeyResult(double StringMs, double IntMs, long Ops);

    private static KeyResult ProbeKeyHashing()
    {
        // Reproduce the engine's tuple-key traffic on a real ConcurrentDictionary of each key type, with the
        // SAME entry count (600k distinct) and the SAME op mix (1.0M TryGetValue over a hit/miss sequence, plus
        // an Add+Remove churn per miss on an "evaluating" set). Populate first, then replay.
        var strCache = new ConcurrentDictionary<(string, string), ComputedValue>();
        var intCache = new ConcurrentDictionary<(int, int), ComputedValue>();
        var strEval = new HashSet<(string, string)>();
        var intEval = new HashSet<(int, int)>();

        // Access sequence mirrors one sweep: for each formula row, self C, then its 2 refs, self D, then its ref.
        // Represent sheet as 0=S,1=Data ; id via (col,row). Strings realistic: "C123456","Data" etc.
        var strSeq = new List<(string, string)>(1_000_000);
        var intSeq = new List<(int, int)>(1_000_000);
        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            // C self
            strSeq.Add((FormulaSheet, "C" + r));
            intSeq.Add((0, EncodeId(2, r)));
            // C -> Data!A{v}, Data!B{v}
            strSeq.Add((DataSheet, "A" + v));
            intSeq.Add((1, EncodeId(0, v)));
            strSeq.Add((DataSheet, "B" + v));
            intSeq.Add((1, EncodeId(1, v)));
            // D self
            strSeq.Add((FormulaSheet, "D" + r));
            intSeq.Add((0, EncodeId(3, r)));
            // D -> Data!B{v}
            strSeq.Add((DataSheet, "B" + v));
            intSeq.Add((1, EncodeId(1, v)));
        }

        var val = ComputedValue.Number(1);

        double bestStr = double.MaxValue,
            bestInt = double.MaxValue;
        for (var trial = 0; trial < 5; trial++)
        {
            strCache.Clear();
            strEval.Clear();
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            double sink = 0;
            for (var i = 0; i < strSeq.Count; i++)
            {
                var k = strSeq[i];
                if (strCache.TryGetValue(k, out var got))
                {
                    if (got.TryGetNumber(out var d))
                        sink += d;
                }
                else
                {
                    strEval.Add(k);
                    strCache[k] = val;
                    strEval.Remove(k);
                    sink += 1;
                }
            }

            sw.Stop();
            GC.KeepAlive(sink);
            bestStr = Math.Min(bestStr, sw.Elapsed.TotalMilliseconds);
        }

        for (var trial = 0; trial < 5; trial++)
        {
            intCache.Clear();
            intEval.Clear();
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            double sink = 0;
            for (var i = 0; i < intSeq.Count; i++)
            {
                var k = intSeq[i];
                if (intCache.TryGetValue(k, out var got))
                {
                    if (got.TryGetNumber(out var d))
                        sink += d;
                }
                else
                {
                    intEval.Add(k);
                    intCache[k] = val;
                    intEval.Remove(k);
                    sink += 1;
                }
            }

            sw.Stop();
            GC.KeepAlive(sink);
            bestInt = Math.Min(bestInt, sw.Elapsed.TotalMilliseconds);
        }

        return new KeyResult(
            bestStr,
            bestInt,
            strSeq.Count + 3L * ((long)FormulaRows * 2 + DataRows * 2)
        );
    }

    private static int EncodeId(int col, int row) => (col << 20) | row;

    // ------------------------------------------------------------ cache-structure headroom

    private readonly record struct CacheStructResult(
        double CDictStr,
        double CDictInt,
        double DictInt,
        double FlatArray
    );

    private static CacheStructResult ProbeCacheStructure()
    {
        // Build the pure-cache access sequence (1.0M lookups) and the distinct insert set (600k). No evaluating
        // churn here — that is measured separately. Sheet 0=S id (col,row); numeric flat index = distinct slot.
        var strSeq = new (string, string)[1_000_000];
        var intSeq = new (int, int)[1_000_000];
        var flatSeq = new int[1_000_000]; // dense slot index, -1 not used here (all pre-touched below)
        var strKeys = new (string, string)[600_000];
        var intKeys = new (int, int)[600_000];

        // Assign a dense slot per distinct cell. S cells: C{r}->2*(r-1), D{r}->2*(r-1)+1  (400k). Data cells:
        // A{v}-> 400000 + (v-1), B{v} -> 500000 + (v-1)  (200k). Total 600k.
        int Slot(int sheet, int col, int idx) =>
            sheet == 0
                ? (col == 2 ? 2 * (idx - 1) : 2 * (idx - 1) + 1)
                : (col == 0 ? 400_000 + (idx - 1) : 500_000 + (idx - 1));

        var w = 0;
        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            strSeq[w] = (FormulaSheet, "C" + r);
            intSeq[w] = (0, EncodeId(2, r));
            flatSeq[w] = Slot(0, 2, r);
            w++;
            strSeq[w] = (DataSheet, "A" + v);
            intSeq[w] = (1, EncodeId(0, v));
            flatSeq[w] = Slot(1, 0, v);
            w++;
            strSeq[w] = (DataSheet, "B" + v);
            intSeq[w] = (1, EncodeId(1, v));
            flatSeq[w] = Slot(1, 1, v);
            w++;
            strSeq[w] = (FormulaSheet, "D" + r);
            intSeq[w] = (0, EncodeId(3, r));
            flatSeq[w] = Slot(0, 3, r);
            w++;
            strSeq[w] = (DataSheet, "B" + v);
            intSeq[w] = (1, EncodeId(1, v));
            flatSeq[w] = Slot(1, 1, v);
            w++;
        }

        var kw = 0;
        for (var r = 1; r <= FormulaRows; r++)
        {
            strKeys[kw] = (FormulaSheet, "C" + r);
            intKeys[kw] = (0, EncodeId(2, r));
            kw++;
            strKeys[kw] = (FormulaSheet, "D" + r);
            intKeys[kw] = (0, EncodeId(3, r));
            kw++;
        }

        for (var v = 1; v <= DataRows; v++)
        {
            strKeys[kw] = (DataSheet, "A" + v);
            intKeys[kw] = (1, EncodeId(0, v));
            kw++;
            strKeys[kw] = (DataSheet, "B" + v);
            intKeys[kw] = (1, EncodeId(1, v));
            kw++;
        }

        var val = ComputedValue.Number(1);

        // Each structure: TryGetValue over the sequence; a MISS inserts the distinct key (first time seen). To
        // mirror the engine (cell computed once), pre-seed nothing and let first touch insert. We approximate by
        // running the sequence against a pre-populated store (all hits) PLUS a separate populate pass timed in —
        // simplest faithful proxy: populate once (600k inserts) then run 1.0M lookups; report populate+lookup.
        double CDictStrRun()
        {
            var best = double.MaxValue;
            for (var t = 0; t < 5; t++)
            {
                var d = new ConcurrentDictionary<(string, string), ComputedValue>();
                GcQuiesce();
                var sw = Stopwatch.StartNew();
                for (var i = 0; i < strKeys.Length; i++)
                    d[strKeys[i]] = val;
                double sink = 0;
                for (var i = 0; i < strSeq.Length; i++)
                {
                    if (d.TryGetValue(strSeq[i], out var g) && g.TryGetNumber(out var x))
                        sink += x;
                }

                sw.Stop();
                GC.KeepAlive(sink);
                best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
            }

            return best;
        }

        double CDictIntRun()
        {
            var best = double.MaxValue;
            for (var t = 0; t < 5; t++)
            {
                var d = new ConcurrentDictionary<(int, int), ComputedValue>();
                GcQuiesce();
                var sw = Stopwatch.StartNew();
                for (var i = 0; i < intKeys.Length; i++)
                    d[intKeys[i]] = val;
                double sink = 0;
                for (var i = 0; i < intSeq.Length; i++)
                {
                    if (d.TryGetValue(intSeq[i], out var g) && g.TryGetNumber(out var x))
                        sink += x;
                }

                sw.Stop();
                GC.KeepAlive(sink);
                best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
            }

            return best;
        }

        double DictIntRun()
        {
            var best = double.MaxValue;
            for (var t = 0; t < 5; t++)
            {
                var d = new Dictionary<(int, int), ComputedValue>(600_000);
                GcQuiesce();
                var sw = Stopwatch.StartNew();
                for (var i = 0; i < intKeys.Length; i++)
                    d[intKeys[i]] = val;
                double sink = 0;
                for (var i = 0; i < intSeq.Length; i++)
                {
                    if (d.TryGetValue(intSeq[i], out var g) && g.TryGetNumber(out var x))
                        sink += x;
                }

                sw.Stop();
                GC.KeepAlive(sink);
                best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
            }

            return best;
        }

        double FlatRun()
        {
            var best = double.MaxValue;
            for (var t = 0; t < 5; t++)
            {
                var arr = new ComputedValue[600_000];
                var has = new bool[600_000];
                GcQuiesce();
                var sw = Stopwatch.StartNew();
                for (var i = 0; i < flatSeq.Length; i++)
                {
                    var slot = flatSeq[i];
                    if (has[slot])
                    {
                        if (arr[slot].TryGetNumber(out _)) { }
                    }
                    else
                    {
                        arr[slot] = val;
                        has[slot] = true;
                    }
                }

                double sink = 0;
                for (var i = 0; i < flatSeq.Length; i++)
                {
                    if (arr[flatSeq[i]].TryGetNumber(out var x))
                        sink += x;
                }

                sw.Stop();
                GC.KeepAlive(sink);
                best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
            }

            return best;
        }

        return new CacheStructResult(CDictStrRun(), CDictIntRun(), DictIntRun(), FlatRun());
    }

    private static TwoMs ProbeEvaluatingChurn()
    {
        var strKeys = new (string, string)[600_000];
        var intKeys = new (int, int)[600_000];
        var kw = 0;
        for (var r = 1; r <= FormulaRows; r++)
        {
            strKeys[kw] = (FormulaSheet, "C" + r);
            intKeys[kw] = (0, EncodeId(2, r));
            kw++;
            strKeys[kw] = (FormulaSheet, "D" + r);
            intKeys[kw] = (0, EncodeId(3, r));
            kw++;
        }

        double bestStr = double.MaxValue,
            bestInt = double.MaxValue;
        for (var t = 0; t < 5; t++)
        {
            var set = new HashSet<(string, string)>();
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            for (var i = 0; i < strKeys.Length; i++)
            {
                set.Add(strKeys[i]);
                set.Remove(strKeys[i]);
            }
            sw.Stop();
            bestStr = Math.Min(bestStr, sw.Elapsed.TotalMilliseconds);
        }

        for (var t = 0; t < 5; t++)
        {
            var set = new HashSet<(int, int)>();
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            for (var i = 0; i < intKeys.Length; i++)
            {
                set.Add(intKeys[i]);
                set.Remove(intKeys[i]);
            }
            sw.Stop();
            bestInt = Math.Min(bestInt, sw.Elapsed.TotalMilliseconds);
        }

        return new TwoMs(bestStr, bestInt);
    }

    // ------------------------------------------------------------ sheet resolution (by name vs handle)

    private readonly record struct SheetResResult(double ByNameMs, double ByHandleMs);

    private static SheetResResult ProbeSheetResolution()
    {
        // 600k resolutions (one per miss). By name: ConcurrentDictionary<string,Sheet> OrdinalIgnoreCase, the
        // real Sheets container type. By handle: a direct array index (what a resolved sheet handle costs).
        var byName = new ConcurrentDictionary<string, Sheet>(StringComparer.OrdinalIgnoreCase);
        var data = new Sheet { Name = DataSheet, Index = 0 };
        var s = new Sheet { Name = FormulaSheet, Index = 1 };
        byName[DataSheet] = data;
        byName[FormulaSheet] = s;
        var handles = new[] { s, data };

        // Sequence: the 600k misses resolve S (400k) and Data (200k) in the sweep's proportion.
        var names = new string[600_000];
        var idxs = new int[600_000];
        var w = 0;
        for (var r = 1; r <= FormulaRows; r++)
        {
            names[w] = FormulaSheet;
            idxs[w] = 0;
            w++; // C self miss
            names[w] = DataSheet;
            idxs[w] = 1;
            w++; // one Data miss amortized
            names[w] = FormulaSheet;
            idxs[w] = 0;
            w++; // D self miss
            if (w >= names.Length)
                break;
        }

        double bestName = double.MaxValue,
            bestHandle = double.MaxValue;
        for (var trial = 0; trial < 5; trial++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            var acc = 0;
            for (var i = 0; i < names.Length; i++)
            {
                if (byName.TryGetValue(names[i], out var sh))
                    acc += sh.Index;
            }

            sw.Stop();
            GC.KeepAlive(acc);
            bestName = Math.Min(bestName, sw.Elapsed.TotalMilliseconds);
        }

        for (var trial = 0; trial < 5; trial++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            var acc = 0;
            for (var i = 0; i < idxs.Length; i++)
            {
                acc += handles[idxs[i]].Index;
            }

            sw.Stop();
            GC.KeepAlive(acc);
            bestHandle = Math.Min(bestHandle, sw.Elapsed.TotalMilliseconds);
        }

        return new SheetResResult(bestName, bestHandle);
    }

    // ------------------------------------------------------------ id lookup in _cells (string vs numeric)

    private readonly record struct TwoMs(double StringMs, double IntMs);

    private static TwoMs ProbeIdLookup()
    {
        // 600k id lookups (one per miss). String: Dictionary<string,Expression> ordinal (the real _cells type).
        // Numeric: a flat array indexed by encoded id (what numeric addressing would allow).
        var strCells = new Dictionary<string, Expression>(600_000);
        for (var r = 1; r <= FormulaRows; r++)
        {
            strCells["C" + r] = BlankValue.Instance;
            strCells["D" + r] = BlankValue.Instance;
        }

        for (var v = 1; v <= DataRows; v++)
        {
            strCells["A" + v] = Show;
            strCells["B" + v] = Hide;
        }

        var ids = new string[600_000];
        var w = 0;
        for (var r = 1; r <= FormulaRows && w < ids.Length - 2; r++)
        {
            var v = (r % DataRows) + 1;
            ids[w++] = "C" + r;
            ids[w++] = "A" + v;
            ids[w++] = "B" + v;
        }

        var flat = new Expression[600_000];
        Array.Fill(flat, BlankValue.Instance);
        var idxSeq = new int[600_000];
        for (var i = 0; i < idxSeq.Length; i++)
            idxSeq[i] = i;

        double bestStr = double.MaxValue,
            bestInt = double.MaxValue;
        for (var trial = 0; trial < 5; trial++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            var acc = 0;
            for (var i = 0; i < ids.Length; i++)
            {
                if (strCells.TryGetValue(ids[i], out var e))
                    acc += e is null ? 0 : 1;
            }

            sw.Stop();
            GC.KeepAlive(acc);
            bestStr = Math.Min(bestStr, sw.Elapsed.TotalMilliseconds);
        }

        for (var trial = 0; trial < 5; trial++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            var acc = 0;
            for (var i = 0; i < idxSeq.Length; i++)
            {
                var e = flat[idxSeq[i]];
                acc += e is null ? 0 : 1;
            }

            sw.Stop();
            GC.KeepAlive(acc);
            bestInt = Math.Min(bestInt, sw.Elapsed.TotalMilliseconds);
        }

        return new TwoMs(bestStr, bestInt);
    }

    // ------------------------------------------------------------ skeleton split

    private readonly record struct SplitResult(double CMs, double DMs);

    private static SplitResult ProbeSkeletonSplit()
    {
        var wbC = BuildData(out var wbCData);
        var sC = wbC.Sheets[FormulaSheet];
        for (var r = 1; r <= FormulaRows * 2; r++)
        {
            var v = (r % DataRows) + 1;
            sC["C" + r] = ExpressionParser.Parse(
                $"=IF(Data!A{v}=\"Show\",Data!B{v}*2,Data!B{v}/2)",
                sC
            );
        }

        GcQuiesce();
        var cMs = SweepSingle(wbC, "C", FormulaRows * 2);

        var wbD = BuildData(out _);
        var sD = wbD.Sheets[FormulaSheet];
        for (var r = 1; r <= FormulaRows * 2; r++)
        {
            var v = (r % DataRows) + 1;
            sD["D" + r] = ExpressionParser.Parse($"=Data!B{v}*2+1", sD);
        }

        GcQuiesce();
        var dMs = SweepSingle(wbD, "D", FormulaRows * 2);

        GC.KeepAlive(wbCData);
        return new SplitResult(cMs, dMs);
    }

    private static double SweepSingle(MySheetWorkbook wb, string prefix, int rows)
    {
        // warm
        for (var r = 1; r <= rows; r++)
        {
            _ = wb.GetCellValue(FormulaSheet, prefix + r);
        }

        var best = double.MaxValue;
        for (var t = 0; t < 3; t++)
        {
            wb.InvalidateCache();
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            double sum = 0;
            for (var r = 1; r <= rows; r++)
            {
                if (wb.GetCellValue(FormulaSheet, prefix + r).TryGetNumber(out var x))
                    sum += x;
            }

            sw.Stop();
            GC.KeepAlive(sum);
            best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
        }

        return best;
    }

    // ------------------------------------------------------------ AST-only (no GetCellValue plumbing)

    private readonly record struct AstResult(double Ms, long Calls);

    private static AstResult ProbeAstOnly()
    {
        // Evaluate the D arithmetic AST directly against a workbook, calling expression.Evaluate on a cached
        // parse tree, bypassing the outer GetCellValue self-lookup — but the inner Data!B ref still goes through
        // GetCellValue (unavoidable, that IS the plumbing). This isolates dispatch+arithmetic+coercion of the
        // node walk from the outer cache/evaluating churn on the self-cell.
        var wb = BuildData(out _);
        var s = wb.Sheets[FormulaSheet];
        var trees = new Expression[FormulaRows];
        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            trees[r - 1] = ExpressionParser.Parse($"=Data!B{v}*2+1", s);
        }

        // warm the Data cache
        for (var v = 1; v <= DataRows; v++)
        {
            _ = wb.GetCellValue(DataSheet, "B" + v);
        }

        var best = double.MaxValue;
        for (var t = 0; t < 3; t++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            double sum = 0;
            for (var r = 0; r < trees.Length; r++)
            {
                var ctx = new EvaluationContext(wb, FormulaSheet, "D" + (r + 1));
                if (trees[r].Evaluate(ctx).TryGetNumber(out var x))
                    sum += x;
            }

            sw.Stop();
            GC.KeepAlive(sum);
            best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
        }

        return new AstResult(best, FormulaRows);
    }

    // ------------------------------------------------------------ builders

    private static MySheetWorkbook BuildFull()
    {
        var wb = BuildData(out _);
        var s = wb.Sheets[FormulaSheet];
        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            s["C" + r] = ExpressionParser.Parse(
                $"=IF(Data!A{v}=\"Show\",Data!B{v}*2,Data!B{v}/2)",
                s
            );
            s["D" + r] = ExpressionParser.Parse($"=Data!B{v}*2+1", s);
        }

        return wb;
    }

    private static MySheetWorkbook BuildData(out Sheet data)
    {
        var wb = new MySheetWorkbook();
        data = wb.Sheets.Add(DataSheet);
        wb.Sheets.Add(FormulaSheet);
        for (var v = 1; v <= DataRows; v++)
        {
            data["A" + v] = v % ShowEvery == 1 ? Show : Hide;
            data["B" + v] = new NumberValue(v);
        }

        return wb;
    }

    private static void GcQuiesce()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}
