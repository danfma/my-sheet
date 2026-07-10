using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Wall-clock harness for the 3.0 LIFETIME structural index (plans/write-time-index-3.0.md, Phase 2).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --structural-index-lifetime
//
// Reproduces the EXACT shape of the user's own benchmark (MYSHEET-PERF-whole-column-scan.md): column A holds a
// fixed 200 populated cells; the rest of the sheet is filled in OTHER columns (so they never touch A:A); each
// iteration does InvalidateCache() then evaluates COUNTIF(A:A,">0").
//
// The Phase-5 (2.x) index was per-EPOCH: dropped by InvalidateCache, so this shape rebuilt it every iteration
// (O(sheet) per epoch, ~5-6× the pre-index NaiveScan). The 3.0 index is write-maintained and LIFETIME-scoped:
// InvalidateCache no longer drops it, so it is built ONCE (the first iteration) and every epoch after serves
// the 200 cells of A:A directly. The mean-per-iteration time should therefore stay ~FLAT as the sheet total
// grows (it tracks column A, ~200 cells — not the whole sheet). That flatness is the Phase-2 deliverable and
// the calibration point for the Phase-3 target table (Aspose ~0.4ms constant).
public static class StructuralIndexLifetimeHarness
{
    private const string DataSheet = "Main";
    private const string CalcSheet = "Calc";
    private const int ColumnACells = 200;

    // Reused by AsposeCompareHarness so both sides share the identical inner iteration count.
    internal const int Iterations = 50;

    // Phase 3 target table: the sheet total sweeps 10k → 500k (the K1 report's axis). Column A stays fixed at
    // 200 cells; the rest fills OTHER columns, so COUNTIF(A:A) always scans only the 200 — the per-epoch time
    // must stay ~flat as the sheet grows (the lifetime index is built once and survives InvalidateCache).
    // Reused by AsposeCompareHarness so both sides sweep the identical sizes.
    internal static readonly int[] SheetSizes = [10_200, 20_200, 40_200, 100_200, 200_200, 500_200];

    public static void Run()
    {
        Console.WriteLine("== 3.0 lifetime structural-index harness (user shape) ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"Column A fixed at {ColumnACells} cells; {Iterations} iterations of "
                + "{{ InvalidateCache(); COUNTIF(A:A,\">0\") }}; mean ms per iteration (best-of-5). "
                + "Expected ~flat across sheet sizes (index built once, survives InvalidateCache)."
        );
        Console.WriteLine();
        Console.WriteLine($"{"Sheet cells", 12} {"Mean ms/iter", 14} {"First-read ms", 14}");

        foreach (var size in SheetSizes)
        {
            var (workbook, calc) = Build(size);
            var firstRead = FirstReadMs(workbook, calc);

            // Rebuild a fresh workbook so the steady-state measurement is not skewed by the warm first read.
            (workbook, calc) = Build(size);
            var mean = MeanPerIteration(workbook, calc);

            Console.WriteLine($"{size, 12:N0} {mean, 14:N3} {firstRead, 14:N3}");
        }
    }

    // The one-shot cost of the FIRST open-range read (which builds the lifetime index once) on a cold workbook.
    private static double FirstReadMs(Workbook workbook, string calcId)
    {
        var stopwatch = Stopwatch.StartNew();
        _ = workbook.GetCellValue(CalcSheet, calcId);
        stopwatch.Stop();
        return stopwatch.Elapsed.TotalMilliseconds;
    }

    // Best-of a few passes of (50 iterations of InvalidateCache + one COUNTIF), reported as mean ms/iteration.
    // Reused by AsposeCompareHarness for the MySheet side of the side-by-side table.
    internal static double MeanPerIteration(Workbook workbook, string calcId)
    {
        var best = double.MaxValue;

        for (var trial = 0; trial < 5; trial++)
        {
            var stopwatch = Stopwatch.StartNew();

            for (var i = 0; i < Iterations; i++)
            {
                workbook.InvalidateCache();
                _ = workbook.GetCellValue(CalcSheet, calcId);
            }

            stopwatch.Stop();
            best = Math.Min(best, stopwatch.Elapsed.TotalMilliseconds / Iterations);
        }

        return best;
    }

    // Column A: 200 populated cells (the data COUNTIF actually scans). The remaining cells fill OTHER columns
    // (1000 rows each), exactly like the user's repro, so they are irrelevant to A:A but inflate the sheet.
    // Reused by AsposeCompareHarness so the MySheet side is built with the exact same shape it compares.
    internal static (Workbook Workbook, string CalcId) Build(int totalCells)
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add(DataSheet);
        var calc = workbook.Sheets.Add(CalcSheet);

        for (var r = 1; r <= ColumnACells; r++)
        {
            data[$"A{r}"] = Number(r);
        }

        var remaining = totalCells - ColumnACells;
        var column = 2; // start at column B; column A stays fixed at 200

        while (remaining > 0)
        {
            var name = ColumnName(column++);
            var rows = Math.Min(remaining, 1_000);

            for (var r = 1; r <= rows; r++)
            {
                data[$"{name}{r}"] = Number(r);
            }

            remaining -= rows;
        }

        const string calcId = "A1";
        calc[calcId] = ExpressionParser.Parse("=COUNTIF(Main!A:A,\">0\")", calc);

        return (workbook, calcId);
    }

    private static string ColumnName(int index)
    {
        var name = string.Empty;

        while (index > 0)
        {
            var remainder = (index - 1) % 26;
            name = (char)('A' + remainder) + name;
            index = (index - 1) / 26;
        }

        return name;
    }

    private static Expressions.Expression Number(double value) =>
        new Expressions.NumberValue(value);
}
