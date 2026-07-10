using System.Diagnostics;
using Aspose.Cells;
using AsposeWorkbook = Aspose.Cells.Workbook;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Side-by-side, SAME-PROCESS, ALL-IN-MEMORY comparison of MySheet 3.0 vs Aspose.Cells on the K1 whole-column
// shape (MYSHEET-PERF-whole-column-scan.md):
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --aspose-compare
//
// Both engines get the IDENTICAL load: column A holds a fixed 200 populated cells; the rest of the sheet is
// filled in OTHER columns (1000 rows each) so they never touch A:A; a Calc cell holds =COUNTIF(Main!A:A,">0").
// The sheet total sweeps 10k -> 500k (the K1 axis). For each size we time the honest per-epoch cost:
//
//   MySheet: InvalidateCache() then GetCellValue(Calc) — reuses StructuralIndexLifetimeHarness verbatim.
//   Aspose:  CalculateFormula() with ForceFullCalculation=true and EnableCalculationChain=false.
//
// WHY THAT APOSE CONFIG IS HONEST (measured, see report): with EnableCalculationChain=true (the default),
// a repeated CalculateFormula() serves the cached result in ~0.000 ms — it recomputes NOTHING. Forcing
// ForceFullCalculation=true and disabling the calc chain makes every CalculateFormula() recompute the whole
// formula set from scratch (measured ~0.21 ms at 500k, vs 0.000 ms cached). That is the honest analogue of
// MySheet's InvalidateCache()+recompute: neither side is allowed to serve a stale cache.
//
// NO save/load on EITHER side (the K1 lesson: a fair comparison keeps both engines fully in memory). Because
// nothing is ever saved, the Aspose evaluation watermark sheet is never injected and the in-memory build has
// no row/cell limit — verified at 500,200 populated cells with the COUNTIF returning the correct 200.
//
// A license is loaded ONLY if ASPOSE_LICENSE_PATH points at an existing file (future-proofing); absent it,
// the evaluation mode is fully sufficient for this in-memory scenario.
public static class AsposeCompareHarness
{
    private const string DataSheet = "Main";
    private const string CalcSheet = "Calc";
    private const int ColumnACells = 200;
    private const int RowsPerColumn = 1_000;

    public static void Run()
    {
        var licenseNote = TryLoadLicense();

        Console.WriteLine(
            "== MySheet 3.0 vs Aspose.Cells — whole-column COUNTIF (K1 shape, in-memory) =="
        );
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. Aspose.Cells {CellsHelper.GetVersion()} ({licenseNote}). "
                + $"Column A fixed at {ColumnACells} cells; {StructuralIndexLifetimeHarness.Iterations} iterations per pass "
                + "of {{ invalidate/full-recalc; COUNTIF(A:A,\">0\") }}; mean ms per iteration (best-of-5). "
                + "Both engines fully in memory, no save/load; Aspose forced to full-recalc every epoch (no cache)."
        );
        Console.WriteLine();
        Console.WriteLine(
            $"{"Sheet cells", 12} {"MySheet ms/iter", 18} {"Aspose ms/iter", 16} {"Aspose/MySheet", 15}"
        );

        foreach (var size in StructuralIndexLifetimeHarness.SheetSizes)
        {
            var (workbook, calc) = StructuralIndexLifetimeHarness.Build(size);
            var mysheet = StructuralIndexLifetimeHarness.MeanPerIteration(workbook, calc);

            var aspose = AsposeMeanPerIteration(size);

            var ratio = mysheet > 0 ? aspose / mysheet : double.NaN;
            Console.WriteLine($"{size, 12:N0} {mysheet, 18:N3} {aspose, 16:N3} {ratio, 15:N2}");
        }
    }

    // Load a license from ASPOSE_LICENSE_PATH if the file exists; otherwise stay in evaluation mode.
    private static string TryLoadLicense()
    {
        var path = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH");

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return "evaluation mode";
        }

        try
        {
            new License().SetLicense(path);
            return $"licensed via {path}";
        }
        catch (Exception ex)
        {
            return $"license load FAILED ({ex.GetType().Name}); evaluation mode";
        }
    }

    // Best-of-5 passes of (50 iterations of a forced FULL recalc), reported as mean ms/iteration — the same
    // shape StructuralIndexLifetimeHarness.MeanPerIteration uses for MySheet, so the two columns are comparable.
    private static double AsposeMeanPerIteration(int totalCells)
    {
        var workbook = BuildAspose(totalCells);
        var calc = workbook.Worksheets[1];

        // Warm the first (index-building) recalc so the steady state is not skewed by it.
        workbook.CalculateFormula();

        var best = double.MaxValue;

        for (var trial = 0; trial < 5; trial++)
        {
            var stopwatch = Stopwatch.StartNew();

            for (var i = 0; i < StructuralIndexLifetimeHarness.Iterations; i++)
            {
                workbook.CalculateFormula();
            }

            stopwatch.Stop();
            best = Math.Min(
                best,
                stopwatch.Elapsed.TotalMilliseconds / StructuralIndexLifetimeHarness.Iterations
            );
        }

        // Sanity: the COUNTIF must actually see the 200 column-A cells (guards a silently-empty build).
        if (Convert.ToInt32(calc.Cells["A1"].Value) != ColumnACells)
        {
            throw new InvalidOperationException(
                $"Aspose COUNTIF returned {calc.Cells["A1"].Value}, expected {ColumnACells} — build shape diverged."
            );
        }

        return best;
    }

    // The Aspose mirror of StructuralIndexLifetimeHarness.Build: column A = 200 numbers, the rest fills OTHER
    // columns (1000 rows each) up to the same total, and a Calc cell holds =COUNTIF(Main!A:A,">0").
    private static AsposeWorkbook BuildAspose(int totalCells)
    {
        var workbook = new AsposeWorkbook();

        // Force every CalculateFormula() to recompute from scratch — no cached-chain shortcut (see class note).
        workbook.Settings.FormulaSettings.ForceFullCalculation = true;
        workbook.Settings.FormulaSettings.EnableCalculationChain = false;

        var data = workbook.Worksheets[0];
        data.Name = DataSheet;
        var calc = workbook.Worksheets.Add(CalcSheet);

        for (var r = 1; r <= ColumnACells; r++)
        {
            data.Cells[$"A{r}"].PutValue((double)r);
        }

        var remaining = totalCells - ColumnACells;
        var column = 1; // 0-based column index; start at column B (1). Column A (0) stays fixed at 200.

        while (remaining > 0)
        {
            var rows = Math.Min(remaining, RowsPerColumn);

            for (var r = 0; r < rows; r++)
            {
                data.Cells[r, column].PutValue((double)(r + 1));
            }

            column++;
            remaining -= rows;
        }

        calc.Cells["A1"].Formula = "=COUNTIF(Main!A:A,\">0\")";

        return workbook;
    }
}
