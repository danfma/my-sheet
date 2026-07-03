using System.Diagnostics;
using Danfma.MySheet;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Wall-clock harness for the whole-column scale scenario (plans/whole-column-performance.md, Phase 0).
// BenchmarkDotNet is great for the REDUCED scale (WholeColumnScaleBenchmarks), but the FULL scale
// baseline is O(F*N) ≈ minutes-to-an-hour — too slow to iterate under BDN. So this measures a single
// wall-clock pass, and for the full scale it SAMPLES a small formula block and extrapolates linearly
// (the cost is linear in F for a fixed N), which is documented in the output.
//
// Run: dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale [--full]
public static class WholeColumnScaleHarness
{
    private const int ReducedDataCells = 50_000;
    private const int ReducedFormulas = 10_000;

    private const int FullDataCells = 500_000;
    private const int FullFormulas = 100_000;
    private const int FullSampleFormulas = 1_000; // measured, then extrapolated ×(FullFormulas / this)

    public static void Run(string[] args)
    {
        var full = args.Contains("--full");

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
            RunFullSampled();
        }
        else
        {
            Console.WriteLine();
            Console.WriteLine("(pass --full to also run the 500k×100k sampled extrapolation)");
        }
    }

    private static void RunReduced()
    {
        Console.WriteLine($"-- REDUCED: {ReducedDataCells:N0} data cells × {ReducedFormulas:N0} formulas --");
        Console.WriteLine($"{"Formula",-14} {"BigColumn (ms)",16} {"NarrowColumn (ms)",18}");

        foreach (var formula in Enum.GetValues<ScaleFormula>())
        {
            var big = TimeOnce(ReducedDataCells, ReducedFormulas, formula, ScaleTarget.BigColumn);
            var narrow = TimeOnce(ReducedDataCells, ReducedFormulas, formula, ScaleTarget.NarrowColumn);
            Console.WriteLine($"{formula,-14} {big.Milliseconds,16:N1} {narrow.Milliseconds,18:N1}");
        }
    }

    private static void RunFullSampled()
    {
        var factor = FullFormulas / (double)FullSampleFormulas;
        Console.WriteLine(
            $"-- FULL (SAMPLED): {FullDataCells:N0} data cells × {FullSampleFormulas:N0} sampled formulas, "
                + $"extrapolated ×{factor:N0} → {FullFormulas:N0} formulas --"
        );
        Console.WriteLine(
            $"{"Formula",-14} {"Big sample (ms)",16} {"Big est. (s)",14} {"Narrow sample (ms)",20} {"Narrow est. (s)",16}"
        );

        foreach (var formula in Enum.GetValues<ScaleFormula>())
        {
            var big = TimeOnce(FullDataCells, FullSampleFormulas, formula, ScaleTarget.BigColumn);
            var narrow = TimeOnce(FullDataCells, FullSampleFormulas, formula, ScaleTarget.NarrowColumn);
            Console.WriteLine(
                $"{formula,-14} {big.Milliseconds,16:N1} {big.Milliseconds * factor / 1000.0,14:N1} "
                    + $"{narrow.Milliseconds,20:N1} {narrow.Milliseconds * factor / 1000.0,16:N1}"
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
        var (workbook, formulaIds) = WholeColumnScaleData.Build(dataCells, formulaCount, formula, target);

        // The structural index / value cache both start cold; drop them so the pass is a true from-cold
        // measurement (this is the load-once-read-once cycle the user reported).
        workbook.InvalidateCache();

        var stopwatch = Stopwatch.StartNew();
        var checksum = WholeColumnScaleData.EvaluateAll(workbook, formulaIds);
        stopwatch.Stop();

        return (stopwatch.Elapsed.TotalMilliseconds, checksum);
    }
}
