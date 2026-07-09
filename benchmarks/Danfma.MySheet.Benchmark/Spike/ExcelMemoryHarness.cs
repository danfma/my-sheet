using System.Diagnostics;
using Aspose.Cells;
using Danfma.MySheet;
using Danfma.MySheet.Excel;
using AsposeWorkbook = Aspose.Cells.Workbook;
using MySheetWorkbook = Danfma.MySheet.Workbook;

namespace Danfma.MySheet.Benchmark.Spike;

// TEMPORARY: measures managed-memory cost (allocations + peak live heap) of the .xlsx write paths, to
// compare our OpenXML integration against Aspose.Cells on the same large file (the K1 sample, 566k cells).
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --excel-memory
public static class ExcelMemoryHarness
{
    public static void Run()
    {
        var samples = FindSamples();
        var myxl = Path.Combine(samples, "k1.myxl");
        if (!File.Exists(myxl))
        {
            Console.WriteLine($"k1.myxl not found at {myxl}");
            return;
        }

        var tmp = Path.Combine(Path.GetTempPath(), "excel-mem");
        Directory.CreateDirectory(tmp);
        var bigXlsx = Path.Combine(tmp, "k1-source.xlsx");

        Console.WriteLine("Loading k1.myxl into a MySheet Workbook...");
        var workbook = MySheetWorkbook.Load(myxl);

        // Scenario E — our ExcelExport (build a fresh .xlsx from scratch). Also produces the input file the
        // load-based scenarios reuse.
        Measure("MySheet ExcelExport (fresh save)", () =>
        {
            workbook.SaveAsExcel(bigXlsx, new ExcelExportOptions { FormulaMode = FormulaMode.ValuesOnly });
        });
        Console.WriteLine($"  -> produced {new FileInfo(bigXlsx).Length / 1024 / 1024} MB xlsx");

        // Scenario A — Aspose load + save (the user's baseline: "open in Excel, save with Aspose").
        var asposeOut = Path.Combine(tmp, "aspose-out.xlsx");
        Measure("Aspose load + save", () =>
        {
            var wb = new AsposeWorkbook(bigXlsx);
            wb.Save(asposeOut);
        });

        // Scenario M — our ExcelMerge into a copy of the same file (loads the existing DOM, writes values).
        var mergeTarget = Path.Combine(tmp, "merge-target.xlsx");
        File.Copy(bigXlsx, mergeTarget, overwrite: true);
        Measure("MySheet ExcelMerge (into existing)", () =>
        {
            workbook.MergeIntoExcel(mergeTarget);
        });
        // The merge wrote the same computed values the export did, so the merged file must match the export.
        VerifyRoundTrip(bigXlsx, mergeTarget);

        Console.WriteLine("\nDone. (Aspose evaluation mode may cap/altér output; memory profile is still indicative.)");
    }

    // Confirms the rewritten file still opens and matches the original on a sample of cells and structure.
    private static void VerifyRoundTrip(string original, string rewritten)
    {
        var a = new AsposeWorkbook(original);
        var b = new AsposeWorkbook(rewritten);

        var okSheets = a.Worksheets.Count == b.Worksheets.Count;
        var mismatches = 0;
        var sampled = 0;

        for (var s = 0; s < Math.Min(a.Worksheets.Count, b.Worksheets.Count); s++)
        {
            var sa = a.Worksheets[s];
            var sb = b.Worksheets[s];
            var maxRow = Math.Min(sa.Cells.MaxDataRow, 200);
            var maxCol = Math.Min(sa.Cells.MaxDataColumn, 40);

            for (var r = 0; r <= maxRow; r++)
            for (var c = 0; c <= maxCol; c++)
            {
                var va = sa.Cells[r, c].StringValue;
                var vb = sb.Cells[r, c].StringValue;
                sampled++;
                if (va != vb) mismatches++;
            }
        }

        Console.WriteLine(
            $"  round-trip: sheets {(okSheets ? "OK" : "MISMATCH")} ({a.Worksheets.Count} vs {b.Worksheets.Count}); "
                + $"sampled {sampled} cells, {mismatches} mismatch(es)"
        );
    }

    private static void Measure(string label, Action action)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var peak = GC.GetTotalMemory(true);
        var stop = false;
        var sampler = new Thread(() =>
        {
            while (!Volatile.Read(ref stop))
            {
                var now = GC.GetTotalMemory(false);
                if (now > peak) peak = now;
                Thread.Sleep(2);
            }
        }) { IsBackground = true };

        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var sw = Stopwatch.StartNew();
        sampler.Start();

        action();

        sw.Stop();
        Volatile.Write(ref stop, true);
        sampler.Join();
        var allocAfter = GC.GetTotalAllocatedBytes(precise: true);

        var liveAfter = GC.GetTotalMemory(true);

        Console.WriteLine(
            $"{label,-40}  time {sw.Elapsed.TotalSeconds,6:N2}s  "
                + $"allocated {Mb(allocAfter - allocBefore),8}  "
                + $"peak-live {Mb(peak),8}  "
                + $"retained {Mb(liveAfter),8}"
        );
    }

    private static string Mb(long bytes) => $"{bytes / 1024.0 / 1024.0:N0} MB";

    private static string FindSamples()
    {
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 10 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "samples");
            if (Directory.Exists(candidate))
            {
                return candidate;
            }

            dir = Path.GetDirectoryName(dir.TrimEnd(Path.DirectorySeparatorChar));
        }

        return Path.Combine(Directory.GetCurrentDirectory(), "samples");
    }
}
