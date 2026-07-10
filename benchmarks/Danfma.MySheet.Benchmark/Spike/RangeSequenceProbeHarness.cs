using System.Diagnostics;
using System.Runtime.CompilerServices;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike;

// INVESTIGATION ONLY (owner question 2026-07-04): with the dense paged value store now holding computed values
// CONTIGUOUSLY in per-column ComputedValue[1024] pages, can the RANGE read paths (RangeReference.
// ExpandComputedValues -> one Workbook.GetCellValue per cell, feeding NumericAggregation.Fold /
// RangeSnapshot.Build / SUMPRODUCT / the *IF(S) criteria engine) expose the page SEGMENTS directly —
// ReadOnlySequence<ComputedValue> (multi-segment over pages), ReadOnlySpan<ComputedValue> per page via a
// visitor, or is IEnumerable good enough? This probe MEASURES the headroom of each strategy so the answer is a
// number, not an opinion. Nothing in Danfma.MySheet is touched.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --range-sequence-probe
//
// WHAT THE CURRENT PATH ACTUALLY DOES (measured/grounded below), per cell of a range:
//   RangeReference.ExpandComputedValues:  id = new CellAddress(col,row).ToId()   // StringBuilder + string ALLOC
//                                         yield return workbook.GetCellValue(SheetName, id)
//   Workbook.GetCellValue(name, id):      CellAddress.TryGetColumnRow(id, ...)   // re-parse the string we built
//                                         store.HandleFor(name)                  // per-cell dict lookup
//                                         store.TryGetDense(handle, col, row, ...)// seqlock read of the page slot
//   => a (col,row) -> string -> (col,row) ROUND TRIP per cell, plus a per-cell sheet-handle dict lookup and an
//      iterator state machine — all of it on top of the actual page-slot read. The dense store made the SLOT
//      read cheap; the range enumerator still addresses every slot by rebuilding and re-parsing its A1 string.
//
// STRATEGIES MEASURED (all fold SUM over the same warmed range; equivalence asserted):
//   V0  real ExpandComputedValues     the true production path, on a real warmed Workbook (grounding).
//   A   string round-trip (replica)   reproduce V0's per-cell work on a standalone real SheetValueStore
//                                      (ToId + TryGetColumnRow + per-cell HandleFor + TryGetDense) — must ~= V0.
//   B   numeric per-cell (replica)    hoist HandleFor once; loop (col,row); TryGetDense(handle,col,row). Kills
//                                      the ToId alloc, the re-parse and the per-cell dict lookup. STILL per cell,
//                                      STILL evaluation-on-demand safe (a miss can fall back to GetCellValue).
//   C   page-span visitor (replica)   iterate the column's pages; fold over ReadOnlySpan<ComputedValue> guarded
//                                      by the presence bitmap. Contiguous, no per-cell dispatch. Needs the cells
//                                      PRESENT (obstacle a) and single-thread/version-checked (obstacle b).
//   D   ReadOnlySequence (replica)    build a ReadOnlySequence<ComputedValue> from custom segments (one per
//                                      page) and fold via SequenceReader-style span walk — the "real" multi-
//                                      segment API, to price the segment-object allocation vs the visitor.
//
// PLUS: RangeSnapshot.Build (materialize ComputedValue[] once) current vs block-copy of pages; a MULTI-COLUMN
// rectangle (obstacle d: pages are vertical per column); and a HALF-PRESENT range (obstacle a: page-span must
// fall back to on-demand evaluation for absent slots, so the win evaporates when the range is not materialized).
//
// METHODOLOGY: best-of-N (min), N=7; GC.Collect between trials; alloc via GC.GetTotalAllocatedBytes(true). The
// replica page store mirrors the production SheetValueStore layout exactly (ComputedValue[1024] per page + a
// 128-byte presence bitmap); A on the replica is cross-checked against V0 on the real engine to prove it tracks.
public static class RangeSequenceProbeHarness
{
    private const string DataSheet = "Data";
    private const int Trials = 7;

    private const int PageShift = 10;
    private const int PageRows = 1 << PageShift; // 1024
    private const int PageMask = PageRows - 1;
    private const int PresenceWords = PageRows / 64;

    public static void Run(string[] args)
    {
        Console.WriteLine(
            "== Range read strategies over the dense paged store (owner question 2026-07-04) =="
        );
        Console.WriteLine(
            $"Runtime {Environment.Version}. sizeof(ComputedValue) = {Unsafe.SizeOf<ComputedValue>()} bytes. "
                + $"Best-of-{Trials} (min); GC.Collect between trials; alloc = GetTotalAllocatedBytes(true)."
        );
        Console.WriteLine();

        SingleColumn();
        Rectangle();
        HalfPresent();
        SnapshotBuild();

        Console.WriteLine(
            "Done. All numbers are min-of-7 on a warmed (fully-present unless noted) range."
        );
    }

    // ============================================================ single dense column A1:A100000
    private static void SingleColumn()
    {
        const int rows = 100_000;
        Console.WriteLine(
            $"-- SINGLE COLUMN  Data!A1:A{rows}  ({rows:N0} contiguous numeric cells, all present) --"
        );

        // Grounding: a REAL warmed Workbook, the true ExpandComputedValues fold.
        var wb = BuildWarmColumn(rows);
        var range = new RangeReference("A1", "A" + rows, DataSheet);
        var ctx = new EvaluationContext(wb, DataSheet, "Z1");
        var v0 = Measure(() => FoldReal(range, ctx), out var sum0);

        // Replica: a standalone real SheetValueStore populated identically, plus a page-layout replica for spans.
        var store = new SheetValueStore();
        var handle = store.HandleFor(DataSheet);
        for (var r = 1; r <= rows; r++)
        {
            store.SetDense(handle, 1, r, ComputedValue.Number(r), tainted: false);
        }

        var replica = new ReplicaColumn(rows);
        for (var r = 1; r <= rows; r++)
        {
            replica.Set(r, ComputedValue.Number(r));
        }

        var a = Measure(() => FoldStringRoundTrip(store, 1, rows), out var sumA);
        var b = Measure(() => FoldNumericPerCell(store, handle, 1, rows), out var sumB);
        var c = Measure(() => FoldPageSpan(replica, rows), out var sumC);
        var d = Measure(() => FoldReadOnlySequence(replica, rows), out var sumD);

        AssertSame("single-column", sum0, sumA, sumB, sumC, sumD);
        Table(
            ("V0 real ExpandComputedValues (grounding)", v0),
            ("A  string round-trip  (real store)", a),
            ("B  numeric per-cell   (real store)", b),
            ("C  page-span visitor  (replica)", c),
            ("D  ReadOnlySequence   (replica)", d)
        );

        var gainB = v0.Ms > 0 ? (a.Ms - b.Ms) / a.Ms * 100 : 0;
        var gainC = a.Ms > 0 ? (a.Ms - c.Ms) / a.Ms * 100 : 0;
        Console.WriteLine(
            $"   A->B  (kill string round-trip)   : {a.Ms - b.Ms, 7:N2} ms faster, {a.Mb - b.Mb, 6:N2} MB less  ({gainB:N0}% of A)"
        );
        Console.WriteLine(
            $"   A->C  (page-span over B)         : {a.Ms - c.Ms, 7:N2} ms faster                ({gainC:N0}% of A)"
        );
        Console.WriteLine(
            $"   B->C  (span vs numeric per-cell) : {b.Ms - c.Ms, 7:N2} ms faster (the marginal span win)"
        );
        Console.WriteLine(
            $"   C->D  (sequence-obj overhead)    : {d.Ms - c.Ms, 7:N2} ms, {d.Mb - c.Mb, 6:N2} MB (segment objects)"
        );
        Console.WriteLine();
    }

    // ============================================================ rectangle 10 cols x 10000 rows (obstacle d)
    private static void Rectangle()
    {
        const int cols = 10;
        const int rowsPer = 10_000;
        Console.WriteLine(
            $"-- RECTANGLE  {cols} cols x {rowsPer:N0} rows = {cols * rowsPer:N0} cells (pages are vertical per column) --"
        );

        var store = new SheetValueStore();
        var handle = store.HandleFor(DataSheet);
        var replicas = new ReplicaColumn[cols + 1];
        for (var col = 1; col <= cols; col++)
        {
            replicas[col] = new ReplicaColumn(rowsPer);
            for (var r = 1; r <= rowsPer; r++)
            {
                var val = ComputedValue.Number(col * 1_000_000 + r);
                store.SetDense(handle, col, r, val, tainted: false);
                replicas[col].Set(r, val);
            }
        }

        // Current: enumerate the rectangle column-by-column, per-cell string round trip (ExpandComputedValues
        // walks column outer, row inner — the real order).
        var a = Measure(() => FoldRectStringRoundTrip(store, cols, rowsPer), out var sumA);
        var b = Measure(() => FoldRectNumericPerCell(store, handle, cols, rowsPer), out var sumB);
        // Page-span per column: a ReadOnlySequence per column works (obstacle d answer — segment the rectangle
        // as a sequence of column runs; each column's pages are its segments).
        var c = Measure(() => FoldRectPageSpan(replicas, cols, rowsPer), out var sumC);

        AssertSame("rectangle", sumA, sumB, sumC);
        Table(
            ("A  string round-trip  (real store)", a),
            ("B  numeric per-cell   (real store)", b),
            ("C  page-span per column (replica)", c)
        );
        Console.WriteLine(
            "   -> a rectangle is a SEQUENCE of column runs; each column's pages are contiguous segments,"
        );
        Console.WriteLine(
            "      so a per-column span visitor (or a ReadOnlySequence spanning one column) is the fit."
        );
        Console.WriteLine();
    }

    // ============================================================ half-present range (obstacle a)
    private static void HalfPresent()
    {
        const int rows = 100_000;
        Console.WriteLine(
            $"-- HALF-PRESENT  Data!A1:A{rows} with only EVEN rows computed (odd rows NOT present) --"
        );
        Console.WriteLine(
            "   The span sees the page but the presence bitmap says a slot is empty: it must be EVALUATED"
        );
        Console.WriteLine(
            "   on demand (the current path's GetCellValue does exactly this). A raw span cannot."
        );

        var replica = new ReplicaColumn(rows);
        var present = 0;
        for (var r = 2; r <= rows; r += 2)
        {
            replica.Set(r, ComputedValue.Number(r));
            present++;
        }

        // Page-span WITH a per-absent-slot fallback (simulating an on-demand evaluate of ~free cost) — this is
        // the honest cost when the range is NOT fully materialized: the span still visits every slot and the
        // fallback pays per hole. The win over per-cell shrinks with presence density.
        var c = Measure(() => FoldPageSpanWithHoles(replica, rows, out _), out _);
        Console.WriteLine(
            $"   present {present:N0}/{rows:N0} (50%). page-span+presence-check fold: {c.Ms:N2} ms, {c.Mb:N2} MB"
        );
        Console.WriteLine(
            "   >>> When a range is NOT fully present the span still has to branch per slot and something"
        );
        Console.WriteLine(
            "       must evaluate the holes; a raw block-copy/SIMD fold is only valid on a PROVEN-full page."
        );
        Console.WriteLine(
            "       Fully-materialized cases: 2nd+ read of the same range in an epoch; a pure DATA range"
        );
        Console.WriteLine(
            "       (plain numbers, present after first touch); post-RangeSnapshot.Build (already an array)."
        );
        Console.WriteLine();
    }

    // ============================================================ RangeSnapshot.Build materialization
    private static void SnapshotBuild()
    {
        const int rows = 100_000;
        Console.WriteLine(
            $"-- RangeSnapshot.Build materialize ComputedValue[{rows}] (Layer-2, the once-per-epoch cost) --"
        );

        var wb = BuildWarmColumn(rows);
        var range = new RangeReference("A1", "A" + rows, DataSheet);
        var ctx = new EvaluationContext(wb, DataSheet, "Z1");

        var replica = new ReplicaColumn(rows);
        for (var r = 1; r <= rows; r++)
        {
            replica.Set(r, ComputedValue.Number(r));
        }

        // Current: List<ComputedValue> fed by ExpandComputedValues, then ToArray (exactly RangeSnapshot.Build).
        var cur = Measure(
            () =>
            {
                var list = new List<ComputedValue>();
                foreach (var v in range.ExpandComputedValues(ctx))
                {
                    list.Add(v);
                }

                var arr = list.ToArray();
                return arr.Length;
            },
            out _
        );

        // Block-copy: pre-sized array, Array.Copy each page's contiguous slice (valid only fully-present).
        var blk = Measure(() => replica.CopyTo(rows).Length, out _);

        // The REAL production RangeSnapshot.Build on the same warmed range — the definitive before/after of the
        // shipped code (before Phase 1: List + ToArray ~= "cur"; after: pre-sized array + block-copy ~= "blk").
        var real = Measure(() => RangeSnapshot.Build(range, ctx).Count, out _);

        Table(
            ("Build current (List + ExpandComputedValues + ToArray)", cur),
            ("Build block-copy pages -> ComputedValue[] (full only)", blk),
            ("Build REAL RangeSnapshot.Build (shipped code)", real)
        );
        Console.WriteLine(
            $"   -> block-copy is {(blk.Ms > 0 ? cur.Ms / blk.Ms : 0):N1}x faster and {cur.Mb - blk.Mb:N1} MB less on a full range;"
        );
        Console.WriteLine(
            "      it is the single clearest win, but only when every slot in the covered pages is present."
        );
        Console.WriteLine();
    }

    // ============================================================ fold implementations

    private static double FoldReal(RangeReference range, EvaluationContext ctx)
    {
        double sum = 0;
        foreach (var v in range.ExpandComputedValues(ctx))
        {
            if (v.TryGetNumber(out var x))
            {
                sum += x;
            }
        }

        return sum;
    }

    // Faithful reproduction of ExpandComputedValues + GetCellValue's per-cell work on a real SheetValueStore.
    private static double FoldStringRoundTrip(SheetValueStore store, int col, int rows)
    {
        double sum = 0;
        for (var r = 1; r <= rows; r++)
        {
            var id = new CellAddress(col, r).ToId(); // ExpandComputedValues: StringBuilder alloc
            CellAddress.TryGetColumnRow(id, out var c, out var rr); // GetCellValue: re-parse
            var h = store.HandleFor(DataSheet); // GetCellValue: per-cell dict lookup
            if (store.TryGetDense(h, c, rr, out var v) && v.TryGetNumber(out var x))
            {
                sum += x;
            }
        }

        return sum;
    }

    private static double FoldNumericPerCell(SheetValueStore store, int handle, int col, int rows)
    {
        double sum = 0;
        for (var r = 1; r <= rows; r++)
        {
            if (store.TryGetDense(handle, col, r, out var v) && v.TryGetNumber(out var x))
            {
                sum += x;
            }
        }

        return sum;
    }

    private static double FoldPageSpan(ReplicaColumn column, int rows)
    {
        double sum = 0;
        var pageCount = column.PageCount(rows);
        for (var pi = 0; pi < pageCount; pi++)
        {
            if (!column.GetPage(pi, rows, out var span, out var present))
            {
                continue;
            }

            for (var i = 0; i < span.Length; i++)
            {
                if ((present[i >> 6] & (1UL << (i & 63))) != 0 && span[i].TryGetNumber(out var x))
                {
                    sum += x;
                }
            }
        }

        return sum;
    }

    private static double FoldReadOnlySequence(ReplicaColumn column, int rows)
    {
        var seq = column.AsSequence(rows);
        double sum = 0;
        var pos = seq.Start;
        while (seq.TryGet(ref pos, out var mem))
        {
            var span = mem.Span;
            for (var i = 0; i < span.Length; i++)
            {
                if (span[i].TryGetNumber(out var x))
                {
                    sum += x;
                }
            }
        }

        return sum;
    }

    private static double FoldPageSpanWithHoles(ReplicaColumn column, int rows, out int holes)
    {
        double sum = 0;
        var missing = 0;
        var pageCount = column.PageCount(rows);
        for (var pi = 0; pi < pageCount; pi++)
        {
            if (!column.GetPage(pi, rows, out var span, out var present))
            {
                // whole page absent -> every row here is a hole a real path would have to evaluate
                missing +=
                    pi == 0
                        ? rows >= PageMask
                            ? PageMask
                            : rows
                        : Math.Min(PageRows, rows - (pi << PageShift));
                continue;
            }

            for (var i = pi == 0 ? 1 : 0; i < span.Length; i++)
            {
                if ((present[i >> 6] & (1UL << (i & 63))) != 0)
                {
                    if (span[i].TryGetNumber(out var x))
                    {
                        sum += x;
                    }
                }
                else
                {
                    missing++; // a real path would GetCellValue here to evaluate the hole
                }
            }
        }

        holes = missing;
        return sum;
    }

    private static double FoldRectStringRoundTrip(SheetValueStore store, int cols, int rowsPer)
    {
        double sum = 0;
        for (var col = 1; col <= cols; col++)
        {
            for (var r = 1; r <= rowsPer; r++)
            {
                var id = new CellAddress(col, r).ToId();
                CellAddress.TryGetColumnRow(id, out var c, out var rr);
                var h = store.HandleFor(DataSheet);
                if (store.TryGetDense(h, c, rr, out var v) && v.TryGetNumber(out var x))
                {
                    sum += x;
                }
            }
        }

        return sum;
    }

    private static double FoldRectNumericPerCell(
        SheetValueStore store,
        int handle,
        int cols,
        int rowsPer
    )
    {
        double sum = 0;
        for (var col = 1; col <= cols; col++)
        {
            for (var r = 1; r <= rowsPer; r++)
            {
                if (store.TryGetDense(handle, col, r, out var v) && v.TryGetNumber(out var x))
                {
                    sum += x;
                }
            }
        }

        return sum;
    }

    private static double FoldRectPageSpan(ReplicaColumn[] replicas, int cols, int rowsPer)
    {
        double sum = 0;
        for (var col = 1; col <= cols; col++)
        {
            var column = replicas[col];
            var pageCount = column.PageCount(rowsPer);
            for (var pi = 0; pi < pageCount; pi++)
            {
                if (!column.GetPage(pi, rowsPer, out var span, out var present))
                {
                    continue;
                }

                for (var i = 0; i < span.Length; i++)
                {
                    if (
                        (present[i >> 6] & (1UL << (i & 63))) != 0
                        && span[i].TryGetNumber(out var x)
                    )
                    {
                        sum += x;
                    }
                }
            }
        }

        return sum;
    }

    // ============================================================ real warmed workbook
    private static Workbook BuildWarmColumn(int rows)
    {
        var wb = new Workbook();
        var data = wb.Sheets.Add(DataSheet);
        for (var r = 1; r <= rows; r++)
        {
            data["A" + r] = new NumberValue(r);
        }

        // Warm: force every cell into the dense store so the range read hits the present path.
        for (var r = 1; r <= rows; r++)
        {
            wb.GetCellValue(DataSheet, "A" + r);
        }

        return wb;
    }

    // ============================================================ measurement
    private readonly record struct M(double Ms, double Mb);

    private static M Measure(Func<double> body, out double result)
    {
        var bestMs = double.MaxValue;
        long bestChurn = long.MaxValue;
        double last = 0;
        for (var t = 0; t < Trials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            last = body();
            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            bestMs = Math.Min(bestMs, sw.Elapsed.TotalMilliseconds);
            bestChurn = Math.Min(bestChurn, churn);
        }

        result = last;
        return new M(bestMs, bestChurn / 1048576d);
    }

    private static M Measure(Func<int> body, out int result)
    {
        var bestMs = double.MaxValue;
        long bestChurn = long.MaxValue;
        var last = 0;
        for (var t = 0; t < Trials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            last = body();
            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            bestMs = Math.Min(bestMs, sw.Elapsed.TotalMilliseconds);
            bestChurn = Math.Min(bestChurn, churn);
        }

        result = last;
        return new M(bestMs, bestChurn / 1048576d);
    }

    private static void Table(params (string Label, M M)[] rows)
    {
        Console.WriteLine($"   {"strategy", -52}{"ms", 8}{"MB churn", 12}");
        Console.WriteLine("   " + new string('-', 72));
        foreach (var (label, m) in rows)
        {
            Console.WriteLine($"   {label, -52}{m.Ms, 8:N2}{m.Mb, 12:N2}");
        }
    }

    private static void AssertSame(string tag, params double[] sums)
    {
        for (var i = 1; i < sums.Length; i++)
        {
            if (Math.Abs(sums[i] - sums[0]) > 1e-6)
            {
                throw new InvalidOperationException(
                    $"[{tag}] fold diverged: variant {i} sum {sums[i]:R} != {sums[0]:R}. Numbers would lie."
                );
            }
        }
    }

    private static void GcQuiesce()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    // ============================================================ replica column (production page layout)
    // Mirrors SheetValueStore.Column/Page: ComputedValue[1024] per page + 128-byte presence bitmap. Used for the
    // span/sequence/block-copy strategies (the production Page is a private nested type, unreachable from here).
    private sealed class ReplicaColumn
    {
        private readonly ReplicaPage?[] _pages;

        public ReplicaColumn(int maxRow)
        {
            _pages = new ReplicaPage?[(maxRow >> PageShift) + 1];
        }

        public void Set(int row, ComputedValue value)
        {
            var pi = row >> PageShift;
            var page = _pages[pi] ??= new ReplicaPage();
            var slot = row & PageMask;
            page.Values[slot] = value;
            page.Present[slot >> 6] |= 1UL << (slot & 63);
        }

        // Number of page indices covering [1..rows] (some may be absent).
        public int PageCount(int rows) => (rows >> PageShift) + 1;

        // Closure-free page accessor: the present page's slot span (0..lastSlot) + presence bitmap. Slot index
        // equals the in-page row offset; slot 0 of page 0 is row 0 (nonexistent) and never present, so a fold
        // over the full span with the presence check is correct. Returns false when the page is absent.
        public bool GetPage(
            int pageIndex,
            int rows,
            out ReadOnlySpan<ComputedValue> span,
            out ulong[] present
        )
        {
            var page = _pages[pageIndex];
            if (page is null)
            {
                span = default;
                present = System.Array.Empty<ulong>();
                return false;
            }

            var baseRow = pageIndex << PageShift;
            var lastSlot = Math.Min(PageMask, rows - baseRow);
            span = page.Values.AsSpan(0, lastSlot + 1);
            present = page.Present;
            return true;
        }

        // Block-copy every present page's contiguous slice into one array (RangeSnapshot.Build alternative).
        public ComputedValue[] CopyTo(int rows)
        {
            var result = new ComputedValue[rows];
            var w = 0;
            var lastPage = rows >> PageShift;
            for (var pi = 0; pi <= lastPage; pi++)
            {
                var page = _pages[pi];
                var baseRow = pi << PageShift;
                var lastSlot = Math.Min(PageMask, rows - baseRow);
                var start = pi == 0 ? 1 : 0;
                var len = lastSlot - start + 1;
                if (page is not null)
                {
                    Array.Copy(page.Values, start, result, w, len);
                }

                w += len;
            }

            return result;
        }

        // A ReadOnlySequence<ComputedValue> made of one segment per present page (the "real" multi-segment API).
        public PageSequence AsSequence(int rows) => new(_pages, rows);
    }

    private sealed class ReplicaPage
    {
        public readonly ComputedValue[] Values = new ComputedValue[PageRows];
        public readonly ulong[] Present = new ulong[PresenceWords];
    }

    // Minimal ReadOnlySequence-shaped walker over pages (prices the per-segment memory/cursor overhead without
    // pulling in System.IO.Pipelines; the real thing would use ReadOnlySequenceSegment<ComputedValue>).
    private readonly struct PageSequence(ReplicaPage?[] pages, int rows)
    {
        public int Start => 0;

        public bool TryGet(ref int pageIndex, out ReadOnlyMemory<ComputedValue> mem)
        {
            var lastPage = rows >> PageShift;
            while (pageIndex <= lastPage)
            {
                var pi = pageIndex++;
                var page = pages[pi];
                if (page is null)
                {
                    continue;
                }

                var baseRow = pi << PageShift;
                var lastSlot = Math.Min(PageMask, rows - baseRow);
                var start = pi == 0 ? 1 : 0;
                mem = new ReadOnlyMemory<ComputedValue>(page.Values, start, lastSlot - start + 1);
                return true;
            }

            mem = default;
            return false;
        }
    }
}
