using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Write-cost gate for the 3.0 write-maintained structural index (plans/write-time-index-3.0.md, Phase 3).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --write-cost
//
// The 3.0 index is maintained ON THE WRITE (the SetCell choke point) instead of rebuilt per read epoch, so a
// bulk Fill of a column pays a small per-set maintenance cost ONCE the index is live. This measures that cost
// across four scenarios: {in-order, out-of-order} fill × {no index built (pure Fill), live index (one
// open-range read before the timed fill)}. The SAME test body — public API only — is copied verbatim into a
// standalone probe that references the published `Danfma.MySheet` 2.9.1 package, so the two trees are compared
// like-for-like; the Phase-3 gate is 3.0 overhead ≤ ~5% vs 2.9.1.
//
// Only the `sheet[id] = expr` loop is timed: the ids and the NumberValue expressions are pre-generated and the
// index-priming read (LiveIndex) runs BEFORE the stopwatch, so the measurement isolates the write path — which
// is the only thing that differs between 2.9.1 and 3.0.
public static class WriteCostHarness
{
    private const int Cells = 500_000;
    private const int Runs = 7;

    public static void Run()
    {
        Console.WriteLine("== 3.0 write-cost gate (SetCell choke point) ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"{Cells:N0} sets through the public indexer; best & median of {Runs} runs; ms for the set loop only."
        );
        Console.WriteLine();
        Console.WriteLine(
            "[isolated] pre-generated ids + NumberValues; the loop times only `sheet[id] = expr` (the pure write path)."
        );
        Console.WriteLine($"{"Scenario", -34} {"Best ms", 10} {"Median ms", 11} {"ns/set", 9}");

        Report("InOrder / NoIndex", inOrder: true, liveIndex: false, realistic: false);
        Report("OutOfOrder / NoIndex", inOrder: false, liveIndex: false, realistic: false);
        Report("InOrder / LiveIndex", inOrder: true, liveIndex: true, realistic: false);
        Report("OutOfOrder / LiveIndex", inOrder: false, liveIndex: true, realistic: false);

        Console.WriteLine();
        Console.WriteLine(
            "[realistic] the loop constructs the id string + NumberValue AND sets — the user-visible Fill cost."
        );
        Report("InOrder / NoIndex (realistic)", inOrder: true, liveIndex: false, realistic: true);
        Report("InOrder / LiveIndex (realistic)", inOrder: true, liveIndex: true, realistic: true);
    }

    private static void Report(string label, bool inOrder, bool liveIndex, bool realistic)
    {
        var samples = new double[Runs];

        for (var i = 0; i < Runs; i++)
        {
            samples[i] = TimeFill(inOrder, liveIndex, realistic);
        }

        Array.Sort(samples);
        var best = samples[0];
        var median = samples[Runs / 2];
        Console.WriteLine(
            $"{label, -24} {best, 10:N1} {median, 11:N1} {median * 1_000_000.0 / Cells, 9:N1}"
        );
    }

    // Milliseconds spent in the fill loop. In [isolated] mode the ids and NumberValues are pre-generated and
    // only `sheet[id] = expr` is timed (pure write path). In [realistic] mode the id string and NumberValue are
    // constructed inside the timed loop too (the user-visible Fill cost), so the write-maintenance overhead is
    // seen as a fraction of a real fill instead of a fraction of the near-free bare dictionary write. The
    // LiveIndex priming read always runs before the stopwatch.
    private static double TimeFill(bool inOrder, bool liveIndex, bool realistic)
    {
        var order = new int[Cells];

        for (var i = 0; i < Cells; i++)
        {
            order[i] = i + 1;
        }

        if (!inOrder)
        {
            var rng = new Random(20260703);

            for (var i = Cells - 1; i > 0; i--)
            {
                var j = rng.Next(i + 1);
                (order[i], order[j]) = (order[j], order[i]);
            }
        }

        string[]? ids = null;
        Expression[]? exprs = null;

        if (!realistic)
        {
            ids = new string[Cells];
            exprs = new Expression[Cells];

            for (var i = 0; i < Cells; i++)
            {
                ids[i] = "A" + order[i];
                exprs[i] = new NumberValue(order[i]);
            }
        }

        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");

        if (liveIndex)
        {
            // Prime the lifetime column index with one open-range read BEFORE the timed fill, so every set in
            // the loop maintains a live index — the 3.0 write-time cost this gate bounds. A seed cell in another
            // column keeps column A empty at build time (so the fill exercises the append/dirty paths cleanly).
            sheet["B1"] = new NumberValue(1);
            var calc = workbook.Sheets.Add("Calc");
            calc["A1"] = ExpressionParser.Parse("=COUNTIF(Main!A:A,\">0\")", calc);
            _ = workbook.GetCellValue("Calc", "A1");
        }

        var stopwatch = Stopwatch.StartNew();

        if (realistic)
        {
            for (var i = 0; i < Cells; i++)
            {
                sheet["A" + order[i]] = new NumberValue(order[i]);
            }
        }
        else
        {
            for (var i = 0; i < Cells; i++)
            {
                sheet[ids![i]] = exprs![i];
            }
        }

        stopwatch.Stop();

        GC.KeepAlive(workbook);
        return stopwatch.Elapsed.TotalMilliseconds;
    }
}
