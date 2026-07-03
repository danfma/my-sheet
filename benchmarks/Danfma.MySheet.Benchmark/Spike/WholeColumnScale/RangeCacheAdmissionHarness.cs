using System.Diagnostics;
using System.Globalization;
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Wall-clock harness for the SECOND-USE ADMISSION regression (plans/whole-column-performance.md, Phase 4).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --range-cache-admission
//
// The Phase 2 cache admits a range on its FIRST read. That is a pure win when a range is read MANY times per
// epoch (the whole-column scenario: 10k formulas over A:A), but a pure LOSS when a range that clears the
// 256-cell threshold is read exactly ONCE per epoch — the formula pays an O(N) materialization (plus a
// derived accelerator it never reuses) on top of the linear scan it would have done anyway. The production
// report (1.5×–6× regressions on small/medium sheets after 2.6.2) is exactly this shape.
//
// Each scenario is measured with the cache DISABLED (the pre-cache linear baseline, via RangeCacheDisabled)
// and ENABLED (the tree's current admission policy). The ratio ENABLED/DISABLED is the regression factor —
// > 1 means the cache is hurting. The three shapes:
//   (a) SlidingWindows  — 10k formulas SUM(A$1:A{500+i}); every range distinct, ≥256, read once.
//   (b) BoundedMatchOnce— 10k formulas MATCH over a distinct ~300-cell window; builds an exact hash for ONE probe.
//   (c) InvalidateLoop  — many epochs of InvalidateCache + a single read of the same big range per epoch.
public static class RangeCacheAdmissionHarness
{
    private const string DataSheet = "Data";
    private const string CalcSheet = "Calc";

    public static void Run()
    {
        Console.WriteLine("== Range-cache second-use admission regression harness ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + "Each scenario: one full pass, cache DISABLED (pre-cache) vs ENABLED (current tree)."
        );
        Console.WriteLine();
        Console.WriteLine($"{"Scenario",-20} {"Disabled (ms)",15} {"Enabled (ms)",14} {"Enabled/Disabled",18}");

        Report("SlidingWindows", BuildSlidingWindows());
        Report("BoundedMatchOnce", BuildBoundedMatchOnce());
        ReportInvalidateLoop();
    }

    private static void Report(string name, (Workbook Workbook, string[] Ids) built)
    {
        // Run each mode a few times and take the best (least noisy) wall-clock.
        var disabled = BestOf(built.Workbook, built.Ids, disabled: true);
        var enabled = BestOf(built.Workbook, built.Ids, disabled: false);
        var ratio = disabled > 0 ? enabled / disabled : 0;
        Console.WriteLine($"{name,-20} {disabled,15:N1} {enabled,14:N1} {ratio,18:N2}x");
    }

    private static double BestOf(Workbook workbook, string[] ids, bool disabled)
    {
        var best = double.MaxValue;

        for (var trial = 0; trial < 5; trial++)
        {
            workbook.RangeCacheDisabled = disabled;
            workbook.InvalidateCache();

            var stopwatch = Stopwatch.StartNew();
            EvaluateAll(workbook, ids);
            stopwatch.Stop();

            best = Math.Min(best, stopwatch.Elapsed.TotalMilliseconds);
        }

        return best;
    }

    // (a) Sliding windows: A1:A{500+i}, each distinct, ≥256, read once.
    private static (Workbook, string[]) BuildSlidingWindows()
    {
        const int formulas = 10_000;
        const int columnCells = 500 + formulas; // 10_500

        var workbook = new Workbook();
        var data = workbook.Sheets.Add(DataSheet);
        var calc = workbook.Sheets.Add(CalcSheet);

        for (var r = 1; r <= columnCells; r++)
        {
            data[$"A{r}"] = Number(r);
        }

        var ids = new string[formulas];
        for (var i = 0; i < formulas; i++)
        {
            var end = 500 + i;
            var id = $"A{i + 1}";
            calc[id] = ExpressionParser.Parse($"=SUM(Data!A$1:A{end})", calc);
            ids[i] = id;
        }

        return (workbook, ids);
    }

    // (b) Bounded distinct MATCH: each formula reads a distinct ~300-cell window once, exact-hash built for ONE probe.
    private static (Workbook, string[]) BuildBoundedMatchOnce()
    {
        const int formulas = 10_000;
        const int window = 300;
        const int columnCells = formulas + window;

        var workbook = new Workbook();
        var data = workbook.Sheets.Add(DataSheet);
        var calc = workbook.Sheets.Add(CalcSheet);

        for (var r = 1; r <= columnCells; r++)
        {
            data[$"A{r}"] = Number(r);
        }

        var ids = new string[formulas];
        for (var i = 0; i < formulas; i++)
        {
            var start = i + 1;
            var end = start + window - 1;
            var probe = (start + end) / 2; // in-range hit
            var id = $"A{i + 1}";
            calc[id] = ExpressionParser.Parse(
                $"=MATCH({probe.ToString(CultureInfo.InvariantCulture)},Data!A{start}:A{end},0)",
                calc
            );
            ids[i] = id;
        }

        return (workbook, ids);
    }

    // (c) InvalidateCache loop: same big range, read once per epoch, many epochs.
    private static void ReportInvalidateLoop()
    {
        const int epochs = 3_000;
        const int columnCells = 11_000;

        var workbook = new Workbook();
        var data = workbook.Sheets.Add(DataSheet);
        var calc = workbook.Sheets.Add(CalcSheet);

        for (var r = 1; r <= columnCells; r++)
        {
            data[$"A{r}"] = Number(r);
        }

        calc["A1"] = ExpressionParser.Parse("=MATCH(5000,Data!A:A,0)", calc);

        var disabled = InvalidateLoopBestOf(workbook, epochs, disabled: true);
        var enabled = InvalidateLoopBestOf(workbook, epochs, disabled: false);
        var ratio = disabled > 0 ? enabled / disabled : 0;
        Console.WriteLine($"{"InvalidateLoop",-20} {disabled,15:N1} {enabled,14:N1} {ratio,18:N2}x");
    }

    private static double InvalidateLoopBestOf(Workbook workbook, int epochs, bool disabled)
    {
        var best = double.MaxValue;

        for (var trial = 0; trial < 3; trial++)
        {
            workbook.RangeCacheDisabled = disabled;

            var stopwatch = Stopwatch.StartNew();
            for (var epoch = 0; epoch < epochs; epoch++)
            {
                workbook.InvalidateCache();
                _ = workbook.GetCellValue(CalcSheet, "A1");
            }

            stopwatch.Stop();
            best = Math.Min(best, stopwatch.Elapsed.TotalMilliseconds);
        }

        return best;
    }

    private static void EvaluateAll(Workbook workbook, string[] ids)
    {
        var checksum = 0.0;
        foreach (var id in ids)
        {
            if (workbook.GetCellValue(CalcSheet, id).AsObject() is double number)
            {
                checksum += number;
            }
        }

        GC.KeepAlive(checksum);
    }

    private static Expressions.Expression Number(double value) => new Expressions.NumberValue(value);
}
