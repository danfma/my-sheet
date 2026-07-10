using System.Diagnostics;
using Danfma.MySheet;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Wall-clock harness for the whole-column scale scenario (plans/whole-column-performance.md, Phases 0 & 3).
// BenchmarkDotNet is great for the REDUCED scale (WholeColumnScaleBenchmarks), but the FULL scale is a
// one-shot manual measurement, so this measures a single wall-clock pass per formula block.
//
//   Reduced only:  dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale
//   + full 100k:   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale --full
//   + sampled trap:dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale --full --sampled
//
// PHASE 3 NOTE — why the FULL block is now MEASURED, not sampled/extrapolated. The Phase 0 baseline (no
// caches) was truly O(F*N): every formula re-scanned N cells, so per-formula cost was constant and a 1k
// sample × 100 was a valid estimate of 100k. WITH the Phase 1+2 caches the cost is O(N + F·log N): the
// snapshot build is a ONE-TIME O(N) cost amortized across the WHOLE block. Sampling 1k formulas then
// multiplying by 100 multiplies that one-time build ×100 and WILDLY over-estimates. So the honest full
// number must be MEASURED over the real 100k formulas — which is exactly what --full now does. --sampled
// still prints the (now-inflated) extrapolation, kept only to document the trap.
public static class WholeColumnScaleHarness
{
    private const int ReducedDataCells = 50_000;
    private const int ReducedFormulas = 10_000;

    private const int FullDataCells = 500_000;
    private const int FullFormulas = 100_000;
    private const int FullSampleFormulas = 1_000; // sampled-mode only; extrapolated ×(FullFormulas / this)

    public static void Run(string[] args)
    {
        var full = args.Contains("--full");
        var sampled = args.Contains("--sampled");

        Console.WriteLine("== Whole-column scale harness ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + "Each row = one InvalidateCache + one full evaluation pass (values memoized)."
        );
        Console.WriteLine();

        RunReduced();

        if (full)
        {
            Console.WriteLine();
            RunFullMeasured();

            if (sampled)
            {
                Console.WriteLine();
                RunFullSampled();
            }
        }
        else
        {
            Console.WriteLine();
            Console.WriteLine("(pass --full to also run the MEASURED 500k×100k full scale)");
        }
    }

    private static void RunReduced()
    {
        Console.WriteLine(
            $"-- REDUCED: {ReducedDataCells:N0} data cells × {ReducedFormulas:N0} formulas --"
        );
        Console.WriteLine($"{"Formula", -14} {"BigColumn (ms)", 16} {"NarrowColumn (ms)", 18}");

        foreach (var formula in Enum.GetValues<ScaleFormula>())
        {
            var big = TimeOnce(ReducedDataCells, ReducedFormulas, formula, ScaleTarget.BigColumn);
            var narrow = TimeOnce(
                ReducedDataCells,
                ReducedFormulas,
                formula,
                ScaleTarget.NarrowColumn
            );
            Console.WriteLine(
                $"{formula, -14} {big.Milliseconds, 16:N1} {narrow.Milliseconds, 18:N1}"
            );
        }
    }

    // The honest full-scale number: MEASURE the real 100k formulas per block. Reports wall-clock and the
    // managed-heap delta the pass RETAINS (cell memo + structural index + range snapshot) as the cache cost,
    // isolated from the workbook itself (which already exists before the timed pass). A total wall-clock and
    // a peak-managed-heap line close the run; wrap the process in `/usr/bin/time -l` for peak RSS.
    private static void RunFullMeasured()
    {
        Console.WriteLine(
            $"-- FULL (MEASURED): {FullDataCells:N0} data cells × {FullFormulas:N0} formulas --"
        );
        Console.WriteLine(
            $"{"Formula", -14} {"Big (ms)", 12} {"Big cache (MB)", 16} {"Narrow (ms)", 14} {"Narrow cache (MB)", 18}"
        );

        var totalMs = 0.0;
        long peakHeap = 0;

        foreach (var formula in Enum.GetValues<ScaleFormula>())
        {
            var big = MeasureWithMemory(
                FullDataCells,
                FullFormulas,
                formula,
                ScaleTarget.BigColumn
            );
            var narrow = MeasureWithMemory(
                FullDataCells,
                FullFormulas,
                formula,
                ScaleTarget.NarrowColumn
            );
            totalMs += big.Milliseconds + narrow.Milliseconds;
            peakHeap = Math.Max(peakHeap, Math.Max(big.PeakHeapBytes, narrow.PeakHeapBytes));

            Console.WriteLine(
                $"{formula, -14} {big.Milliseconds, 12:N1} {big.CacheBytes / 1_048_576.0, 16:N1} "
                    + $"{narrow.Milliseconds, 14:N1} {narrow.CacheBytes / 1_048_576.0, 18:N1}"
            );
        }

        Console.WriteLine();
        Console.WriteLine(
            $"TOTAL measured wall-clock (14 blocks of 100k formulas): {totalMs / 1000.0:N2} s "
                + $"({totalMs:N0} ms). Peak managed heap during any single block: "
                + $"{peakHeap / 1_048_576.0:N1} MB."
        );
    }

    private static void RunFullSampled()
    {
        var factor = FullFormulas / (double)FullSampleFormulas;
        Console.WriteLine(
            $"-- FULL (SAMPLED — OVER-ESTIMATES post-cache, see file header): {FullDataCells:N0} data cells × "
                + $"{FullSampleFormulas:N0} sampled formulas, extrapolated ×{factor:N0} → {FullFormulas:N0} formulas --"
        );
        Console.WriteLine(
            $"{"Formula", -14} {"Big sample (ms)", 16} {"Big est. (s)", 14} {"Narrow sample (ms)", 20} {"Narrow est. (s)", 16}"
        );

        foreach (var formula in Enum.GetValues<ScaleFormula>())
        {
            var big = TimeOnce(FullDataCells, FullSampleFormulas, formula, ScaleTarget.BigColumn);
            var narrow = TimeOnce(
                FullDataCells,
                FullSampleFormulas,
                formula,
                ScaleTarget.NarrowColumn
            );
            Console.WriteLine(
                $"{formula, -14} {big.Milliseconds, 16:N1} {big.Milliseconds * factor / 1000.0, 14:N1} "
                    + $"{narrow.Milliseconds, 20:N1} {narrow.Milliseconds * factor / 1000.0, 16:N1}"
            );
        }
    }

    private static (double Milliseconds, double Checksum) TimeOnce(
        int dataCells,
        int formulaCount,
        ScaleFormula formula,
        ScaleTarget target
    )
    {
        var (workbook, formulaIds) = WholeColumnScaleData.Build(
            dataCells,
            formulaCount,
            formula,
            target
        );

        // The structural index / value cache both start cold; drop them so the pass is a true from-cold
        // measurement (this is the load-once-read-once cycle the user reported).
        workbook.InvalidateCache();

        var stopwatch = Stopwatch.StartNew();
        var checksum = WholeColumnScaleData.EvaluateAll(workbook, formulaIds);
        stopwatch.Stop();

        return (stopwatch.Elapsed.TotalMilliseconds, checksum);
    }

    private readonly record struct MeasuredBlock(
        double Milliseconds,
        long CacheBytes,
        long PeakHeapBytes
    );

    private static MeasuredBlock MeasureWithMemory(
        int dataCells,
        int formulaCount,
        ScaleFormula formula,
        ScaleTarget target
    )
    {
        var (workbook, formulaIds) = WholeColumnScaleData.Build(
            dataCells,
            formulaCount,
            formula,
            target
        );
        workbook.InvalidateCache();

        // Settle the workbook so the delta isolates what the EVALUATION pass allocates and retains (the cell
        // memo + structural index + range snapshot), not the workbook build. A compacting, blocking collection
        // reclaims the large arrays on the LOH (500k-element snapshots) that a plain collection can leave —
        // otherwise the before/after delta picks up cross-block LOH noise.
        SettleHeap();
        var heapBeforeEval = GC.GetTotalMemory(forceFullCollection: true);

        var stopwatch = Stopwatch.StartNew();
        _ = WholeColumnScaleData.EvaluateAll(workbook, formulaIds);
        stopwatch.Stop();

        // Retained heap with the caches still alive (a real full collection, workbook + all caches rooted).
        SettleHeap();
        var heapWithCaches = GC.GetTotalMemory(forceFullCollection: true);
        var cacheBytes = Math.Max(0, heapWithCaches - heapBeforeEval);

        // Drop the block so its footprint does not carry into the next block's peak.
        GC.KeepAlive(workbook);
        return new MeasuredBlock(stopwatch.Elapsed.TotalMilliseconds, cacheBytes, heapWithCaches);
    }

    private static void SettleHeap()
    {
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
    }
}
