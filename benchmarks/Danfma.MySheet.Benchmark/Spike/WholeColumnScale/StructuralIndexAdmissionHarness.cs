using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Wall-clock harness for the STRUCTURAL-INDEX second-use admission (plans/whole-column-performance.md, Phase 5).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --structural-index-admission
//
// Reproduces the EXACT shape of the user's own benchmark (MYSHEET-PERF-whole-column-scan.md): column A holds a
// fixed 200 populated cells; the rest of the sheet is filled in OTHER columns (so they never touch A:A); each
// iteration does InvalidateCache() then evaluates COUNTIF(A:A,">0"). Since Phase 1 made the open-range read
// index-backed UNCONDITIONALLY, every epoch rebuilt the WHOLE-SHEET index (O(N) pass + per-column allocation)
// just to serve 200 cells — ~5-6× slower than the pre-index NaiveScan of 2.6.1.
//
// Three modes measured per sheet size (via the Workbook.StructuralIndexMode lever):
//   ForceNaive  — always NaiveScan: the pre-index (2.6.1) baseline.
//   ForceBuild  — build the index on every read: the pre-Phase-5 tree (reproduces the regression).
//   Admission   — Phase 5 default: NaiveScan the first read, build on the second. With one read per epoch this
//                 never builds → parity with the pre-index baseline.
public static class StructuralIndexAdmissionHarness
{
    private const string DataSheet = "Main";
    private const string CalcSheet = "Calc";
    private const int ColumnACells = 200;
    private const int Iterations = 50;

    private static readonly int[] SheetSizes = [2_200, 10_200, 20_200, 40_200];

    public static void Run()
    {
        Console.WriteLine("== Structural-index second-use admission harness (user shape) ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"Column A fixed at {ColumnACells} cells; {Iterations} iterations of "
                + "{{ InvalidateCache(); COUNTIF(A:A,\">0\") }}; mean ms per iteration."
        );
        Console.WriteLine();
        Console.WriteLine(
            $"{"Sheet cells",12} {"ForceNaive",12} {"ForceBuild",12} {"Admission",12} "
                + $"{"Build/Naive",12} {"Admit/Naive",12}"
        );

        foreach (var size in SheetSizes)
        {
            var (workbook, calc) = Build(size);

            var naive = MeanPerIteration(workbook, calc, StructuralIndexMode.ForceNaive);
            var build = MeanPerIteration(workbook, calc, StructuralIndexMode.ForceBuild);
            var admit = MeanPerIteration(workbook, calc, StructuralIndexMode.Admission);

            var buildRatio = naive > 0 ? build / naive : 0;
            var admitRatio = naive > 0 ? admit / naive : 0;

            Console.WriteLine(
                $"{size,12:N0} {naive,12:N3} {build,12:N3} {admit,12:N3} "
                    + $"{buildRatio,11:N2}x {admitRatio,11:N2}x"
            );
        }
    }

    // Best-of a few passes of (50 iterations of InvalidateCache + one COUNTIF), reported as mean ms/iteration.
    private static double MeanPerIteration(Workbook workbook, string calcId, StructuralIndexMode mode)
    {
        workbook.StructuralIndexMode = mode;

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
    private static (Workbook Workbook, string CalcId) Build(int totalCells)
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

    private static Expressions.Expression Number(double value) => new Expressions.NumberValue(value);
}
