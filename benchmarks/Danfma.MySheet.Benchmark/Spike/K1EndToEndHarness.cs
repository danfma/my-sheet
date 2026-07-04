using System.Diagnostics;
using Aspose.Cells;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using AsposeWorkbook = Aspose.Cells.Workbook;
using MySheetWorkbook = Danfma.MySheet.Workbook;

namespace Danfma.MySheet.Benchmark.Spike;

// Side-by-side, SAME-PROCESS, ALL-IN-MEMORY, PER-PHASE decomposition of MySheet vs Aspose.Cells on the K1
// shape, to localize the "in-memory 4.94s/4.63GB (MySheet) vs 3.25s/3.09GB (Aspose), gap 1.5x" the user's
// Copilot measured (plans/function-coverage-roadmap.md, "Comparativo JUSTO no 3.0.0"):
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --k1-endtoend
//
// LOAD (identical on both sides):
//   sheet "Data": 100k rows x 2 columns. A = alternating text "Show"/"Hide" (~1 in 7 is "Show");
//                 B = numbers (B{v} = v).
//   sheet "S":    400k formulas from 2 skeletons, cross-sheet, with v = (r % 100k) + 1:
//                   C{r} = IF(Data!A{v}="Show", Data!B{v}*2, Data!B{v}/2)
//                   D{r} = Data!B{v}*2+1
//   ~600k cells total (200k Data + 400k S) / 400k formulas — the K1 shape (K1 itself is ~566k cells /
//   396k formulas with the same cross-sheet references).
//
// PHASES (each timed with a Stopwatch and allocation-measured with GC.GetTotalAllocatedBytes(precise:true)):
//   1. build+fill  MySheet: value sets (StringValue/NumberValue via the indexer) + ExpressionParser.Parse of
//                           every formula via the indexer.
//                  Aspose:  PutValue for the values + cell.Formula = "..." (the IDENTICAL formula strings).
//   2. compute     MySheet: a sweep of GetCellValue over every formula cell (first evaluation; populates the
//                           value cache).
//                  Aspose:  CalculateFormula() with ForceFullCalculation=true + EnableCalculationChain=false
//                           (the honest "no cached chain" config from AsposeCompareHarness — see that class).
//   3. extract     read the computed value of EVERY formula cell into an aggregate (sum + count).
//                  MySheet is measured in TWO variants on separate table rows, to quantify the boxing lever
//                  the user's Copilot flagged:
//                    (a) ComputedValue direct — TryGetNumber, no boxing.
//                    (b) via AsObject() — the object? bridge that boxes each numeric scalar.
//                  Aspose:  cell.DoubleValue (the non-boxing numeric accessor; see EXTRACT note below).
//   4. TOTAL       per engine (build + compute + extract(a); the extract(b) delta is reported separately).
//
// METHODOLOGY: single process, best-of-N (min) per phase, N reported below (the whole run is expensive, so N
// is small). Phases run in the SAME order on both engines; GC.Collect between engines and between trials;
// keep the machine quiet. The min time and min allocation are tracked per phase independently (both are
// near-constant across trials; JIT warmup falls out of the min).
//
// SANITY / EQUIVALENCE: the extract aggregate (sum over all formula cells) MUST match between MySheet and
// Aspose within 1e-9 — otherwise the two engines are not computing the same thing and the comparison lies.
// The per-cell values are exact in double (B integer; B*2, B/2, B*2+1 all exact) and the aggregate is summed
// by THIS harness in the identical order on both sides, so the expected difference is 0. The value is printed.
//
// ASYMMETRY (documented, affects the per-phase reading): MySheet parses every formula EAGERLY in build+fill
// (ExpressionParser.Parse builds the AST there). Aspose only STORES the formula string on assignment and
// parses lazily inside CalculateFormula. So MySheet's parse cost lands in phase 1 while Aspose's lands in
// phase 2. Read build vs compute as a PAIR per engine, and the TOTAL as the honest cross-engine number.
//
// EXTRACT accessor choice (Aspose): cell.DoubleValue is used (not cell.Value) so the Aspose row is the
// non-boxing analogue of MySheet variant (a) — cell.Value returns object and would box, muddying the compare.
//
// A license is loaded ONLY if ASPOSE_LICENSE_PATH points at an existing file; absent it, evaluation mode is
// fully sufficient (nothing is ever saved, so no watermark sheet and no in-memory cell cap — same rationale
// as AsposeCompareHarness).
public static class K1EndToEndHarness
{
    private const string DataSheet = "Data";
    private const string FormulaSheet = "S";
    private const int DataRows = 100_000;   // Data!A/B populated rows
    private const int FormulaRows = 200_000; // C{r} and D{r} each -> 400k formulas
    private const int ShowEvery = 7;         // ~1 in 7 A-cells are "Show"
    private const int Trials = 3;            // best-of-N; N reported in the header

    private readonly struct Phase
    {
        public Phase(double ms, long bytes)
        {
            Ms = ms;
            Bytes = bytes;
        }

        public double Ms { get; }
        public long Bytes { get; }
    }

    public static void Run()
    {
        var licenseNote = TryLoadLicense();

        Console.WriteLine("== MySheet vs Aspose.Cells — K1 end-to-end, per-phase, in-memory ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. Aspose.Cells "
                + $"{CellsHelper.GetVersion()} ({licenseNote}). Load: Data {DataRows:N0} rows x 2 cols "
                + $"(A text Show/Hide ~1/{ShowEvery}, B numbers) + S {FormulaRows * 2:N0} formulas (2 skeletons, "
                + "cross-sheet). ~600k cells / 400k formulas. Best-of-" + Trials + " (min) per phase; "
                + "GC.Collect between engines/trials; allocation via GC.GetTotalAllocatedBytes(precise:true).");
        Console.WriteLine();

        // Best (min) time and best (min) alloc per phase, tracked independently across trials.
        var msBuild = new double[2];      // [0]=MySheet [1]=Aspose
        var msCompute = new double[2];
        var msExtractA = new double[2];   // MySheet (a) ComputedValue ; Aspose DoubleValue
        var msExtractB = new double[1];   // MySheet (b) AsObject only
        var byBuild = new long[2];
        var byCompute = new long[2];
        var byExtractA = new long[2];
        var byExtractB = new long[1];

        for (var i = 0; i < 2; i++)
        {
            msBuild[i] = msCompute[i] = msExtractA[i] = double.MaxValue;
            byBuild[i] = byCompute[i] = byExtractA[i] = long.MaxValue;
        }

        msExtractB[0] = double.MaxValue;
        byExtractB[0] = long.MaxValue;

        double mysheetAggregate = double.NaN;
        double asposeAggregate = double.NaN;
        long mysheetCount = 0;
        long asposeCount = 0;

        for (var trial = 0; trial < Trials; trial++)
        {
            // ---- MySheet ----
            GcQuiesce();
            var (mBuild, wb) = MeasureMySheetBuild();
            var mCompute = MeasureMySheetCompute(wb);
            var (mExtractA, mExtractB, aggA, cntA) = MeasureMySheetExtract(wb);

            mysheetAggregate = aggA;
            mysheetCount = cntA;

            msBuild[0] = Math.Min(msBuild[0], mBuild.Ms);
            byBuild[0] = Math.Min(byBuild[0], mBuild.Bytes);
            msCompute[0] = Math.Min(msCompute[0], mCompute.Ms);
            byCompute[0] = Math.Min(byCompute[0], mCompute.Bytes);
            msExtractA[0] = Math.Min(msExtractA[0], mExtractA.Ms);
            byExtractA[0] = Math.Min(byExtractA[0], mExtractA.Bytes);
            msExtractB[0] = Math.Min(msExtractB[0], mExtractB.Ms);
            byExtractB[0] = Math.Min(byExtractB[0], mExtractB.Bytes);

            wb = null;
            GcQuiesce();

            // ---- Aspose ----
            var (aBuild, awb) = MeasureAsposeBuild();
            var aCompute = MeasureAsposeCompute(awb);
            var (aExtract, aggAsp, cntAsp) = MeasureAsposeExtract(awb);

            asposeAggregate = aggAsp;
            asposeCount = cntAsp;

            msBuild[1] = Math.Min(msBuild[1], aBuild.Ms);
            byBuild[1] = Math.Min(byBuild[1], aBuild.Bytes);
            msCompute[1] = Math.Min(msCompute[1], aCompute.Ms);
            byCompute[1] = Math.Min(byCompute[1], aCompute.Bytes);
            msExtractA[1] = Math.Min(msExtractA[1], aExtract.Ms);
            byExtractA[1] = Math.Min(byExtractA[1], aExtract.Bytes);

            awb = null;
            GcQuiesce();
        }

        // Equivalence sanity — the comparison is only honest if both engines computed the same aggregate.
        if (mysheetCount != asposeCount)
        {
            throw new InvalidOperationException(
                $"Formula-cell COUNT diverged: MySheet {mysheetCount} vs Aspose {asposeCount}.");
        }

        if (Math.Abs(mysheetAggregate - asposeAggregate) > 1e-9)
        {
            throw new InvalidOperationException(
                $"Aggregate DIVERGED beyond 1e-9: MySheet {mysheetAggregate:R} vs Aspose {asposeAggregate:R} "
                    + $"(delta {mysheetAggregate - asposeAggregate:R}). The comparison would be lying.");
        }

        // ---- Table ----
        var totalMySheet = msBuild[0] + msCompute[0] + msExtractA[0];
        var totalAspose = msBuild[1] + msCompute[1] + msExtractA[1];
        var totalBytesMySheet = SumClamp(byBuild[0], byCompute[0], byExtractA[0]);
        var totalBytesAspose = SumClamp(byBuild[1], byCompute[1], byExtractA[1]);

        Console.WriteLine($"{"Phase",-26} {"Engine",-10} {"ms (best)",12} {"MB alloc (best)",18}");
        Console.WriteLine(new string('-', 68));
        Row("build+fill", "MySheet", msBuild[0], byBuild[0]);
        Row("build+fill", "Aspose", msBuild[1], byBuild[1]);
        Row("compute", "MySheet", msCompute[0], byCompute[0]);
        Row("compute", "Aspose", msCompute[1], byCompute[1]);
        Row("extract (a) direct", "MySheet", msExtractA[0], byExtractA[0]);
        Row("extract (b) AsObject", "MySheet", msExtractB[0], byExtractB[0]);
        Row("extract DoubleValue", "Aspose", msExtractA[1], byExtractA[1]);
        Console.WriteLine(new string('-', 68));
        Row("TOTAL (build+compute+a)", "MySheet", totalMySheet, totalBytesMySheet);
        Row("TOTAL (build+compute+a)", "Aspose", totalAspose, totalBytesAspose);
        Console.WriteLine();

        var timeGap = totalAspose > 0 ? totalMySheet / totalAspose : double.NaN;
        var allocGap = totalBytesAspose > 0 ? (double)totalBytesMySheet / totalBytesAspose : double.NaN;
        Console.WriteLine($"TOTAL gap (MySheet/Aspose): time {timeGap:N2}x, alloc {allocGap:N2}x.");

        var boxingMs = msExtractB[0] - msExtractA[0];
        var boxingMb = (byExtractB[0] - byExtractA[0]) / (1024d * 1024d);
        Console.WriteLine(
            $"AsObject() boxing lever (MySheet extract b - a): +{boxingMs:N1} ms, +{boxingMb:N1} MB over "
                + $"{mysheetCount:N0} formula cells.");
        Console.WriteLine();
        Console.WriteLine(
            $"Equivalence sanity: aggregate MySheet {mysheetAggregate:R} == Aspose {asposeAggregate:R} "
                + $"(delta {Math.Abs(mysheetAggregate - asposeAggregate):R} <= 1e-9); "
                + $"count {mysheetCount:N0} == {asposeCount:N0}. OK.");
    }

    private static void Row(string phase, string engine, double ms, long bytes)
    {
        var mb = bytes == long.MaxValue ? double.NaN : bytes / (1024d * 1024d);
        Console.WriteLine($"{phase,-26} {engine,-10} {ms,12:N1} {mb,18:N1}");
    }

    private static long SumClamp(params long[] values)
    {
        long sum = 0;

        foreach (var v in values)
        {
            if (v == long.MaxValue)
            {
                return long.MaxValue;
            }

            sum += v;
        }

        return sum;
    }

    // ------------------------------------------------------------------ MySheet

    private static (Phase, MySheetWorkbook) MeasureMySheetBuild()
    {
        var before = GC.GetTotalAllocatedBytes(true);
        var sw = Stopwatch.StartNew();

        var workbook = new MySheetWorkbook();
        var data = workbook.Sheets.Add(DataSheet);
        var s = workbook.Sheets.Add(FormulaSheet);

        for (var v = 1; v <= DataRows; v++)
        {
            data["A" + v] = v % ShowEvery == 1 ? Show : Hide;
            data["B" + v] = new NumberValue(v);
        }

        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            s["C" + r] = ExpressionParser.Parse(
                $"=IF(Data!A{v}=\"Show\",Data!B{v}*2,Data!B{v}/2)", s);
            s["D" + r] = ExpressionParser.Parse($"=Data!B{v}*2+1", s);
        }

        sw.Stop();
        var bytes = GC.GetTotalAllocatedBytes(true) - before;
        return (new Phase(sw.Elapsed.TotalMilliseconds, bytes), workbook);
    }

    private static Phase MeasureMySheetCompute(MySheetWorkbook workbook)
    {
        var before = GC.GetTotalAllocatedBytes(true);
        var sw = Stopwatch.StartNew();

        // First evaluation of every formula cell — populates the value cache. A throwaway aggregate keeps the
        // reads from being elided; the measured "extract" pass below re-reads from the now-warm cache.
        double sink = 0;

        for (var r = 1; r <= FormulaRows; r++)
        {
            if (workbook.GetCellValue(FormulaSheet, "C" + r).TryGetNumber(out var c))
            {
                sink += c;
            }

            if (workbook.GetCellValue(FormulaSheet, "D" + r).TryGetNumber(out var d))
            {
                sink += d;
            }
        }

        sw.Stop();
        var bytes = GC.GetTotalAllocatedBytes(true) - before;
        GC.KeepAlive(sink);
        return new Phase(sw.Elapsed.TotalMilliseconds, bytes);
    }

    // Both extract variants over the warm cache, INTERLEAVED and best-of-R, so the boxing lever is read on an
    // even footing. If (a) and (b) ran once each in fixed order (a then b), the second pass would run faster
    // purely from warmer CPU/branch state and could mask (or invert) the boxing cost. Here both code paths are
    // JIT/cache-warmed once (discarded), then each is timed R times interleaved and the MIN per variant is
    // kept — the boxing cost then lands in the allocation delta (deterministic, captured on one clean pass)
    // and, to whatever extent it exceeds the noise floor, in the time delta.
    //   (a) ComputedValue direct — TryGetNumber, no boxing.
    //   (b) via AsObject() — the object? bridge that boxes each numeric scalar.
    private static (Phase A, Phase B, double Aggregate, long Count) MeasureMySheetExtract(
        MySheetWorkbook workbook)
    {
        const int reps = 5;

        // Warm both code paths once (discarded).
        _ = ExtractDirect(workbook, out _);
        _ = ExtractAsObject(workbook, out _);

        var bestA = double.MaxValue;
        var bestB = double.MaxValue;
        double aggregate = 0;
        long count = 0;

        for (var rep = 0; rep < reps; rep++)
        {
            var swA = Stopwatch.StartNew();
            aggregate = ExtractDirect(workbook, out count);
            swA.Stop();
            bestA = Math.Min(bestA, swA.Elapsed.TotalMilliseconds);

            var swB = Stopwatch.StartNew();
            var aggB = ExtractAsObject(workbook, out var cntB);
            swB.Stop();
            bestB = Math.Min(bestB, swB.Elapsed.TotalMilliseconds);

            if (Math.Abs(aggregate - aggB) > 1e-9 || count != cntB)
            {
                throw new InvalidOperationException(
                    $"MySheet extract variants disagree: direct={aggregate}/{count} vs AsObject={aggB}/{cntB}.");
            }
        }

        // Allocation is deterministic per pass — capture it on one clean pass of each after the timing loop.
        var beforeA = GC.GetTotalAllocatedBytes(true);
        _ = ExtractDirect(workbook, out _);
        var bytesA = GC.GetTotalAllocatedBytes(true) - beforeA;

        var beforeB = GC.GetTotalAllocatedBytes(true);
        _ = ExtractAsObject(workbook, out _);
        var bytesB = GC.GetTotalAllocatedBytes(true) - beforeB;

        return (new Phase(bestA, bytesA), new Phase(bestB, bytesB), aggregate, count);
    }

    private static double ExtractDirect(MySheetWorkbook workbook, out long count)
    {
        double sum = 0;
        long n = 0;

        for (var r = 1; r <= FormulaRows; r++)
        {
            if (workbook.GetCellValue(FormulaSheet, "C" + r).TryGetNumber(out var c))
            {
                sum += c;
                n++;
            }

            if (workbook.GetCellValue(FormulaSheet, "D" + r).TryGetNumber(out var d))
            {
                sum += d;
                n++;
            }
        }

        count = n;
        return sum;
    }

    private static double ExtractAsObject(MySheetWorkbook workbook, out long count)
    {
        double sum = 0;
        long n = 0;

        for (var r = 1; r <= FormulaRows; r++)
        {
            if (workbook.GetCellValue(FormulaSheet, "C" + r).AsObject() is double c)
            {
                sum += c;
                n++;
            }

            if (workbook.GetCellValue(FormulaSheet, "D" + r).AsObject() is double d)
            {
                sum += d;
                n++;
            }
        }

        count = n;
        return sum;
    }

    private static Expression Show { get; } = new StringValue("Show");

    private static Expression Hide { get; } = new StringValue("Hide");

    // ------------------------------------------------------------------ Aspose

    private static (Phase, AsposeWorkbook) MeasureAsposeBuild()
    {
        var before = GC.GetTotalAllocatedBytes(true);
        var sw = Stopwatch.StartNew();

        var workbook = new AsposeWorkbook();

        // Force every CalculateFormula() to recompute from scratch — no cached-chain shortcut (see class note
        // and AsposeCompareHarness). Set before any formula is added.
        workbook.Settings.FormulaSettings.ForceFullCalculation = true;
        workbook.Settings.FormulaSettings.EnableCalculationChain = false;

        var data = workbook.Worksheets[0];
        data.Name = DataSheet;
        var s = workbook.Worksheets.Add(FormulaSheet);

        var dataCells = data.Cells;

        for (var v = 1; v <= DataRows; v++)
        {
            dataCells[v - 1, 0].PutValue(v % ShowEvery == 1 ? "Show" : "Hide"); // A (col 0)
            dataCells[v - 1, 1].PutValue((double)v);                            // B (col 1)
        }

        var sCells = s.Cells;

        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            sCells[r - 1, 2].Formula = $"=IF(Data!A{v}=\"Show\",Data!B{v}*2,Data!B{v}/2)"; // C (col 2)
            sCells[r - 1, 3].Formula = $"=Data!B{v}*2+1";                                  // D (col 3)
        }

        sw.Stop();
        var bytes = GC.GetTotalAllocatedBytes(true) - before;
        return (new Phase(sw.Elapsed.TotalMilliseconds, bytes), workbook);
    }

    private static Phase MeasureAsposeCompute(AsposeWorkbook workbook)
    {
        var before = GC.GetTotalAllocatedBytes(true);
        var sw = Stopwatch.StartNew();

        workbook.CalculateFormula();

        sw.Stop();
        var bytes = GC.GetTotalAllocatedBytes(true) - before;
        return new Phase(sw.Elapsed.TotalMilliseconds, bytes);
    }

    // Aspose extract via cell.DoubleValue — the non-boxing numeric accessor (mirror of MySheet variant a).
    // Warmed once then best-of-R (min), same methodology as the MySheet extract, so the extract rows compare.
    private static (Phase, double Aggregate, long Count) MeasureAsposeExtract(AsposeWorkbook workbook)
    {
        const int reps = 5;
        var cells = workbook.Worksheets[1].Cells;

        _ = AsposeExtract(cells, out _); // warm (discarded)

        var best = double.MaxValue;
        double aggregate = 0;
        long count = 0;

        for (var rep = 0; rep < reps; rep++)
        {
            var sw = Stopwatch.StartNew();
            aggregate = AsposeExtract(cells, out count);
            sw.Stop();
            best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
        }

        var before = GC.GetTotalAllocatedBytes(true);
        _ = AsposeExtract(cells, out _);
        var bytes = GC.GetTotalAllocatedBytes(true) - before;

        return (new Phase(best, bytes), aggregate, count);
    }

    private static double AsposeExtract(Cells cells, out long count)
    {
        double sum = 0;
        long n = 0;

        for (var r = 1; r <= FormulaRows; r++)
        {
            sum += cells[r - 1, 2].DoubleValue; // C
            n++;
            sum += cells[r - 1, 3].DoubleValue; // D
            n++;
        }

        count = n;
        return sum;
    }

    // ------------------------------------------------------------------ shared

    private static void GcQuiesce()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

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
}
