using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Benchmark.Spike;

// Phase 0 SPIKE of the dense value store (plans/dense-value-store-4.0.md). INVESTIGATION ONLY — nothing in
// Danfma.MySheet is touched. Every number here is produced by benchmark-side probes that reproduce the exact
// GetCellValue-shape traffic of the K1 compute sweep against the paged dense store the design proposes, and
// compares it to today's ConcurrentDictionary<(string,string)> cache under the identical load.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --dense-store-spike
//
// WHAT THIS PROVES (or disproves) BEFORE any production change:
//   1. The two concurrency variants of the store (design item 2):
//        (a) concurrent striped — pages published via Interlocked/Volatile, ComputedValue (multi-word struct,
//            torn-write possible) protected by a per-page SEQLOCK (lock-free readers, single-writer via a
//            per-page CAS spinlock). Justification for seqlock over a striped Monitor lock: the traffic is
//            1.0M reads : 600k writes — read-heavy. A striped Monitor PER READ reintroduces exactly the
//            Monitor.Enter cost (~45% of today's compute self-time, plans line 13) that motivates the whole
//            redesign; a seqlock keeps the 1.0M reads lock-free and pays only on the 600k writes. That is the
//            elegant answer to "protect a multi-word struct on a read-heavy path".
//        (b) single-threaded — plain pages, no synchronization (design item 2b: evaluation declared
//            single-threaded per epoch). The array-pure ceiling.
//      Baseline: the ConcurrentDictionary<(string,string), ComputedValue> in production today, same load.
//   2. The on-the-fly DERIVATION tax (design item 5 / VETO i): string id -> (col,row) via
//      CellAddress.TryGetColumnRow (no-alloc, already in the lib) + sheet name -> int handle, over the full
//      1.0M-access traffic. If this tax exceeds 15% of the GAIN (baseline - dense), AST numerization (breaking)
//      enters the discussion; otherwise the store ships as an additive 3.2.0 living behind GetCellValue.
//   3. Allocation: today's cache churns ~162 MB/sweep (plans line 5). Target retained store ~10-20 MB.
//   4. SPARSITY: 10k cells scattered over columns A..ZZ and rows up to 1M — the paged store must not explode
//      (only touched pages allocate). This probe reports the TRUTH about page granularity, including the
//      worst case where cells scatter one-per-page.
//   5. VISITED-BIT: the HashSet<(string,string)> cycle guard (_evaluating) vs a per-slot bit (design item 3).
//
// METHODOLOGY: best-of-N (min) per probe, N>=5 for the pure probes; GC.Collect between trials; allocation via
// GC.GetTotalAllocatedBytes(precise:true); retained memory via GC.GetTotalMemory(forceFullCollection:true)
// while the store is kept alive. The id/name strings are PRE-MATERIALIZED (a real sheet already holds its cell
// ids — GetCellValue never allocates them), so only what each structure itself allocates is measured.
//
// ADDRESSING MODEL (design item 1): per sheet, dense pages of ComputedValue[PageRows] indexed by
// column -> block of rows (PageRows=1024 rows per column). A page is lazily allocated on first touch. Presence
// is an explicit per-slot bit (NOT the ComputedValueKind sentinel: a legitimately-Blank computed cell must be
// distinguishable from an absent one). The per-column page-pointer array is sized from the sheet's known row
// bound — the 3.0 structural index already tracks these bounds, so passing them in is not cheating.
public static class DenseStoreSpikeHarness
{
    private const string DataSheet = "Data";
    private const string FormulaSheet = "S";
    private const int DataRows = 100_000;
    private const int FormulaRows = 200_000; // C{r} and D{r} -> 400k formulas
    private const int PageRows = 1024;
    private const int PageShift = 10; // row >> 10 -> page index (PageRows == 1 << PageShift)
    private const int PageMask = PageRows - 1; // row & 1023 -> slot in page
    private const int PureTrials = 7; // best-of-N for the pure probes (N>=5 required)

    // Sheet handles used by the derivation: 0 = S (formulas), 1 = Data.
    private const int HandleS = 0;
    private const int HandleData = 1;

    public static void Run(string[] args)
    {
        Console.WriteLine("== Dense value store — Phase 0 spike (K1 GetCellValue-shape) ==");
        Console.WriteLine(
            $"Runtime {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"Shape: Data {DataRows:N0}x2 (A,B), S {FormulaRows * 2:N0} formulas (C,D). "
                + "Traffic: 1.00M GetCellValue (600k miss/insert + 400k hit). "
                + $"Page {PageRows} rows/col. Best-of-{PureTrials} (min); GC.Collect between trials."
        );
        Console.WriteLine(
            $"sizeof(ComputedValue) = {Unsafe.SizeOf<ComputedValue>()} bytes (multi-word => torn-write subject)."
        );
        Console.WriteLine();

        var seq = BuildAccessSequence();

        // ------------------------------------------------------------------ derivation tax (VETO i)
        // OWNER DIRECTIVE (2026-07-03): the on-the-fly derivation MUST be span + bit math — zero temp alloc, no
        // Substring/Split/regex/alloc-ToUpper, address packed into a single long, page indexing by shift/mask.
        // We verify CellAddress.TryGetColumnRow already satisfies this (char-indexed single pass, char-overload
        // ToUpperInvariant, no substring/Parse => zero alloc) AND provide a span-based packed-long variant
        // (ParseIdToPacked) to compare head-to-head. The span/packed number is the one that decides the veto.
        var parseLib = ProbeIdParse(seq, useSpan: false); // CellAddress.TryGetColumnRow(string)
        var parseSpan = ProbeIdParse(seq, useSpan: true); // ParseIdToPacked(ReadOnlySpan<char>) -> long
        var nameOnlyMs = ProbeNameResolutionOnly(seq);
        var derivationMs = parseSpan.Ms + nameOnlyMs; // derivation total in the packed-bits format

        // ------------------------------------------------------------------ the three contenders, full shape
        var baseline = ProbeBaselineDictionary(seq);
        var denseB = ProbeDenseStore(seq, concurrent: false, derive: true);
        var denseA = ProbeDenseStore(seq, concurrent: true, derive: true);

        // Isolation: the dense path WITHOUT derivation (pre-derived handle/col/row), to cross-check the tax.
        var denseBpre = ProbeDenseStore(seq, concurrent: false, derive: false);
        var denseApre = ProbeDenseStore(seq, concurrent: true, derive: false);

        // Equivalence sanity — every variant must sum to the same value or the addressing is wrong.
        AssertEquivalence(seq, baseline, denseA, denseB, denseBpre, denseApre);

        // ------------------------------------------------------------------ visited-bit (design item 3)
        var visited = ProbeVisitedGuard();

        // ------------------------------------------------------------------ sparsity (design item 4)
        var sparse = ProbeSparsity();

        // ================================================================== report
        Console.WriteLine(
            "-- derivation tax (on-the-fly string -> (col,row) + name -> handle, full 1.0M traffic) --"
        );
        Console.WriteLine(
            "   OWNER DIRECTIVE: span + bit math, zero temp alloc, packed-long address, shift/mask page index."
        );
        Console.WriteLine(
            $"  id parse — CellAddress.TryGetColumnRow(string) : {parseLib.Ms, 8:N1} ms  ({parseLib.ChurnMb:N2} MB churn)"
        );
        Console.WriteLine(
            $"  id parse — span ParseIdToPacked -> long        : {parseSpan.Ms, 8:N1} ms  ({parseSpan.ChurnMb:N2} MB churn)"
        );
        Console.WriteLine(
            $"  name resolve — ConcurrentDictionary OrdinalIC  : {nameOnlyMs, 8:N1} ms"
        );
        Console.WriteLine(
            $"  derivation total (span parse + name)           : {derivationMs, 8:N1} ms   <- veto (i) number"
        );
        Console.WriteLine(
            $"  (both parsers are zero-alloc single-pass; lib TryGetColumnRow already satisfies the directive)"
        );
        Console.WriteLine();

        Console.WriteLine($"{"Variant", -42}{"ms", 10}{"MB churn", 12}{"MB heap", 12}");
        Console.WriteLine(new string('-', 76));
        Row("baseline ConcurrentDictionary<(str,str)>", baseline);
        Row("dense (a) seqlock  + derivation", denseA);
        Row("dense (b) plain    + derivation", denseB);
        Row("dense (a) seqlock  (pre-derived)", denseApre);
        Row("dense (b) plain    (pre-derived)", denseBpre);
        Console.WriteLine(new string('-', 76));
        Console.WriteLine(
            "  MB churn = GC.GetTotalAllocatedBytes, byte-accurate (this is the plan's target metric:"
        );
        Console.WriteLine(
            "  baseline ~162MB -> target 10-20MB). MB heap = GetTotalMemory upper bound; it over-reports"
        );
        Console.WriteLine(
            "  the store's 24KB page arrays ~2x (GC reservation), so trust churn for the store footprint."
        );
        Console.WriteLine();

        // VETO i — derivation cost vs the gain.
        var gainB = baseline.Ms - denseB.Ms;
        var gainA = baseline.Ms - denseA.Ms;
        var derivShareOfGainB = gainB > 0 ? derivationMs / gainB * 100 : double.NaN;
        var derivShareOfGainA = gainA > 0 ? derivationMs / gainA * 100 : double.NaN;
        // Cross-check the standalone tax against the full-minus-prederived delta.
        var taxDeltaB = denseB.Ms - denseBpre.Ms;
        var taxDeltaA = denseA.Ms - denseApre.Ms;

        Console.WriteLine(
            "== VETO (i) — does on-the-fly derivation eat the gain? (threshold 15%) =="
        );
        Console.WriteLine(
            $"  gain (b) = baseline {baseline.Ms:N1} - dense(b) {denseB.Ms:N1} = {gainB:N1} ms"
        );
        Console.WriteLine(
            $"  gain (a) = baseline {baseline.Ms:N1} - dense(a) {denseA.Ms:N1} = {gainA:N1} ms"
        );
        Console.WriteLine($"  derivation standalone            : {derivationMs:N1} ms");
        Console.WriteLine(
            $"  derivation as %% of gain (b)      : {derivShareOfGainB:N1}%  ({(derivShareOfGainB > 15 ? "OVER 15% -> AST numerization enters" : "under 15% -> additive store OK")})"
        );
        Console.WriteLine($"  derivation as %% of gain (a)      : {derivShareOfGainA:N1}%");
        Console.WriteLine(
            $"  cross-check tax = full - prederived: (b) {taxDeltaB:N1} ms, (a) {taxDeltaA:N1} ms"
        );
        Console.WriteLine();

        // VETO ii — how much does (a) leave on the table vs (b)?
        var aVsB = denseA.Ms - denseB.Ms;
        var aVsBpre = denseApre.Ms - denseBpre.Ms;
        Console.WriteLine("== VETO (ii) — concurrent (a) cost vs single-threaded (b) ceiling ==");
        Console.WriteLine(
            $"  (a) - (b) with derivation   : {aVsB:N1} ms  ({(denseB.Ms > 0 ? aVsB / denseB.Ms * 100 : 0):N1}% over (b))"
        );
        Console.WriteLine(
            $"  (a) - (b) pre-derived (pure): {aVsBpre:N1} ms  ({(denseBpre.Ms > 0 ? aVsBpre / denseBpre.Ms * 100 : 0):N1}% over (b))"
        );
        Console.WriteLine(
            "  -> the seqlock+CAS overhead on 600k writes + version read on 1.0M reads is the price of"
        );
        Console.WriteLine(
            "     keeping evaluation concurrent; (b) is the ceiling if evaluation is declared single-threaded."
        );
        Console.WriteLine();

        // Visited-bit.
        Console.WriteLine("== VISITED-BIT (cycle guard) — 600k enter/exit ==");
        Console.WriteLine(
            $"  HashSet<(string,string)> Add+Remove : {visited.HashSetMs, 8:N1} ms  ({visited.HashSetChurnMb:N1} MB churn)"
        );
        Console.WriteLine(
            $"  per-slot bit  set+clear             : {visited.BitMs, 8:N1} ms  ({visited.BitChurnMb:N1} MB churn)"
        );
        Console.WriteLine();

        // Sparsity.
        Console.WriteLine(
            "== SPARSITY — 10k cells over columns A..ZZ, rows up to 1,000,000 (byte-accurate analytic) =="
        );
        foreach (var s in sparse)
        {
            Console.WriteLine(
                $"  {s.Label, -18}: touched {s.TouchedPages, 5} pages -> dense {s.DenseTotalMb, 7:N2} MB "
                    + $"(pages {s.PageDataMb, 6:N2} + ptr-arrays {s.PointerMb, 5:N2}); baseline dict ~{s.BaselineDictMb:N2} MB"
            );
        }
        Console.WriteLine(
            "  NOTE: page granularity is 1024 ROWS PER COLUMN. Cells clustered within pages stay cheap;"
        );
        Console.WriteLine(
            "  cells scattered one-per-page inflate the dense store (each touched page is ~24 KB of slots)."
        );
        Console.WriteLine(
            "  >>> RISK for Phase 1: pathological row-scatter DOES balloon the store; real sheets cluster."
        );
        Console.WriteLine();
    }

    private static void Row(string label, Measured m)
    {
        Console.WriteLine($"{label, -42}{m.Ms, 10:N1}{m.ChurnMb, 12:N1}{m.RetainedMb, 14:N1}");
    }

    // ============================================================ access sequence
    // 1.0M accesses mirroring one compute sweep (identical to K1ComputeProfileHarness): per formula row r,
    //   C self, Data!A{v}, Data!B{v}, D self, Data!B{v}   with v = (r % DataRows) + 1.
    // 600k distinct cells (400k S: C{r},D{r}; 200k Data: A{v},B{v}) -> 600k first-touch misses, 400k hits.
    private sealed class AccessSequence
    {
        public required string[] Name; // sheet name per access ("S" / "Data")
        public required string[] Id; // cell id per access ("C123", "A45", ...)
        public required int[] Handle; // pre-derived sheet handle
        public required int[] Col; // pre-derived 1-based column
        public required int[] Row; // pre-derived 1-based row
        public required double[] Value; // deterministic value stored/read for that cell
        public required double ExpectedSum; // sum every correct variant must reproduce
        public required int MaxColS;
        public required int MaxColData;
    }

    private static AccessSequence BuildAccessSequence()
    {
        const int n = FormulaRows * 5; // 1.0M
        var name = new string[n];
        var id = new string[n];
        var handle = new int[n];
        var col = new int[n];
        var row = new int[n];
        var value = new double[n];

        var w = 0;
        void Add(string sheetName, int h, string cellId, int c, int r)
        {
            name[w] = sheetName;
            id[w] = cellId;
            handle[w] = h;
            col[w] = c;
            row[w] = r;
            value[w] = CellValue(h, c, r);
            w++;
        }

        for (var r = 1; r <= FormulaRows; r++)
        {
            var v = (r % DataRows) + 1;
            Add(FormulaSheet, HandleS, "C" + r, 3, r); // C = column 3
            Add(DataSheet, HandleData, "A" + v, 1, v); // A = column 1
            Add(DataSheet, HandleData, "B" + v, 2, v); // B = column 2
            Add(FormulaSheet, HandleS, "D" + r, 4, r); // D = column 4
            Add(DataSheet, HandleData, "B" + v, 2, v); // repeat -> hit
        }

        // Expected sum = sum over every access of the value that cell holds (hit reads it back, miss stores it).
        double expected = 0;
        for (var i = 0; i < n; i++)
        {
            expected += value[i];
        }

        return new AccessSequence
        {
            Name = name,
            Id = id,
            Handle = handle,
            Col = col,
            Row = row,
            Value = value,
            ExpectedSum = expected,
            MaxColS = 4,
            MaxColData = 2,
        };
    }

    // Deterministic per-cell value. handle*2e9 + col*1e6 + row stays < 2^53 (row<=1e6), so it is exact in
    // double and a mis-routed slot (wrong sheet/col/row) yields a different number -> the sum diverges.
    private static double CellValue(int handle, int col, int row) =>
        handle * 2_000_000_000d + col * 1_000_000d + row;

    // ============================================================ derivation probes

    // Span-based, zero-allocation id parser: letters -> column, digits -> row, single pass, pure bit/arith, no
    // Substring/Split/ToUpper-alloc. Returns the address PACKED into one long: (col << 32) | (uint)row, or -1.
    // The hot store path unpacks with a shift and a mask; the page index is row >> 10 and the slot is row & 1023.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static long ParseIdToPacked(ReadOnlySpan<char> id)
    {
        long col = 0;
        var i = 0;
        while (i < id.Length)
        {
            var c = id[i];
            int lv;
            if (c >= 'A' && c <= 'Z')
            {
                lv = c - 'A' + 1;
            }
            else if (c >= 'a' && c <= 'z')
            {
                lv = c - 'a' + 1;
            }
            else
            {
                break;
            }

            col = col * 26 + lv;
            i++;
        }

        if (i == 0 || i == id.Length)
        {
            return -1;
        }

        long row = 0;
        for (; i < id.Length; i++)
        {
            var c = id[i];
            if (c is < '0' or > '9')
            {
                return -1;
            }

            row = row * 10 + (c - '0');
        }

        return (col << 32) | (uint)row;
    }

    // Measures ONE parser over the full 1.0M-id traffic: min ms + allocation churn (proves zero-alloc).
    private static Measured ProbeIdParse(AccessSequence seq, bool useSpan)
    {
        double bestMs = double.MaxValue;
        long bestChurn = long.MaxValue;
        for (var t = 0; t < PureTrials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            long acc = 0;
            if (useSpan)
            {
                for (var i = 0; i < seq.Id.Length; i++)
                {
                    var addr = ParseIdToPacked(seq.Id[i]);
                    acc += (int)(addr >> 32) + (int)addr; // col + row
                }
            }
            else
            {
                for (var i = 0; i < seq.Id.Length; i++)
                {
                    if (CellAddress.TryGetColumnRow(seq.Id[i], out var c, out var r))
                    {
                        acc += c + r;
                    }
                }
            }

            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            GC.KeepAlive(acc);
            bestMs = Math.Min(bestMs, sw.Elapsed.TotalMilliseconds);
            bestChurn = Math.Min(bestChurn, churn);
        }

        return new Measured(bestMs, bestChurn / 1048576d, 0, 0);
    }

    private static double ProbeNameResolutionOnly(AccessSequence seq)
    {
        var nameDict = BuildNameDict();
        var best = double.MaxValue;
        for (var t = 0; t < PureTrials; t++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            long acc = 0;
            for (var i = 0; i < seq.Name.Length; i++)
            {
                acc += nameDict[seq.Name[i]];
            }

            sw.Stop();
            GC.KeepAlive(acc);
            best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
        }

        return best;
    }

    private static ConcurrentDictionary<string, int> BuildNameDict()
    {
        var d = new ConcurrentDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        d[FormulaSheet] = HandleS;
        d[DataSheet] = HandleData;
        return d;
    }

    // ============================================================ measured contenders

    private readonly record struct Measured(
        double Ms,
        double ChurnMb,
        double RetainedMb,
        double Sum
    );

    private static Measured ProbeBaselineDictionary(AccessSequence seq)
    {
        double bestMs = double.MaxValue;
        long bestChurn = long.MaxValue;
        double sum = 0;

        for (var t = 0; t < PureTrials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            var cache = new ConcurrentDictionary<(string, string), ComputedValue>();
            double s = 0;
            for (var i = 0; i < seq.Id.Length; i++)
            {
                var key = (seq.Name[i], seq.Id[i]);
                if (cache.TryGetValue(key, out var got))
                {
                    if (got.TryGetNumber(out var x))
                    {
                        s += x; // hit
                    }
                }
                else
                {
                    var v = ComputedValue.Number(seq.Value[i]);
                    cache[key] = v; // miss / insert
                    s += seq.Value[i];
                }
            }

            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            sum = s;
            bestMs = Math.Min(bestMs, sw.Elapsed.TotalMilliseconds);
            bestChurn = Math.Min(bestChurn, churn);
            GC.KeepAlive(cache);
        }

        // Retained: build one, hold it, measure.
        var retained = MeasureRetained(() =>
        {
            var cache = new ConcurrentDictionary<(string, string), ComputedValue>();
            for (var i = 0; i < seq.Id.Length; i++)
            {
                var key = (seq.Name[i], seq.Id[i]);
                if (!cache.ContainsKey(key))
                {
                    cache[key] = ComputedValue.Number(seq.Value[i]);
                }
            }

            return cache;
        });

        return new Measured(bestMs, bestChurn / 1048576d, retained, sum);
    }

    private static Measured ProbeDenseStore(AccessSequence seq, bool concurrent, bool derive)
    {
        var nameDict = BuildNameDict();
        double bestMs = double.MaxValue;
        long bestChurn = long.MaxValue;
        double sum = 0;

        for (var t = 0; t < PureTrials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();

            var store = NewStore(concurrent, seq);
            double s = concurrent
                ? RunSeq(store.Seq!, seq, nameDict, derive)
                : RunPlain(store.Plain!, seq, nameDict, derive);

            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            sum = s;
            bestMs = Math.Min(bestMs, sw.Elapsed.TotalMilliseconds);
            bestChurn = Math.Min(bestChurn, churn);
            GC.KeepAlive(store);
        }

        var retained = MeasureRetained(() =>
        {
            var store = NewStore(concurrent, seq);
            if (concurrent)
            {
                RunSeq(store.Seq!, seq, nameDict, derive: false);
            }
            else
            {
                RunPlain(store.Plain!, seq, nameDict, derive: false);
            }

            return store;
        });

        return new Measured(bestMs, bestChurn / 1048576d, retained, sum);
    }

    private readonly record struct StorePair(SeqStore? Seq, PlainStore? Plain);

    private static StorePair NewStore(bool concurrent, AccessSequence seq)
    {
        // Two sheet stores by handle; column/row bounds from the shape (structural index provides these live).
        if (concurrent)
        {
            var s = new SeqStore(
                new SeqSheetStore(seq.MaxColS, FormulaRows),
                new SeqSheetStore(seq.MaxColData, DataRows)
            );
            return new StorePair(s, null);
        }

        var p = new PlainStore(
            new PlainSheetStore(seq.MaxColS, FormulaRows),
            new PlainSheetStore(seq.MaxColData, DataRows)
        );
        return new StorePair(null, p);
    }

    private static double RunPlain(
        PlainStore store,
        AccessSequence seq,
        ConcurrentDictionary<string, int> nameDict,
        bool derive
    )
    {
        double s = 0;
        for (var i = 0; i < seq.Id.Length; i++)
        {
            int h,
                c,
                r;
            if (derive)
            {
                h = nameDict[seq.Name[i]];
                var addr = ParseIdToPacked(seq.Id[i]); // span + bit pack (owner directive)
                c = (int)(addr >> 32);
                r = (int)addr;
            }
            else
            {
                h = seq.Handle[i];
                c = seq.Col[i];
                r = seq.Row[i];
            }

            var sheet = store.Sheet(h);
            if (sheet.TryGet(c, r, out var val))
            {
                if (val.TryGetNumber(out var x))
                {
                    s += x; // hit
                }
            }
            else
            {
                var v = ComputedValue.Number(seq.Value[i]);
                sheet.Set(c, r, v); // miss
                s += seq.Value[i];
            }
        }

        return s;
    }

    private static double RunSeq(
        SeqStore store,
        AccessSequence seq,
        ConcurrentDictionary<string, int> nameDict,
        bool derive
    )
    {
        double s = 0;
        for (var i = 0; i < seq.Id.Length; i++)
        {
            int h,
                c,
                r;
            if (derive)
            {
                h = nameDict[seq.Name[i]];
                var addr = ParseIdToPacked(seq.Id[i]); // span + bit pack (owner directive)
                c = (int)(addr >> 32);
                r = (int)addr;
            }
            else
            {
                h = seq.Handle[i];
                c = seq.Col[i];
                r = seq.Row[i];
            }

            var sheet = store.Sheet(h);
            if (sheet.TryGet(c, r, out var val))
            {
                if (val.TryGetNumber(out var x))
                {
                    s += x; // hit (lock-free seqlock read)
                }
            }
            else
            {
                var v = ComputedValue.Number(seq.Value[i]);
                sheet.Set(c, r, v); // miss (seqlock write)
                s += seq.Value[i];
            }
        }

        return s;
    }

    // ============================================================ visited-bit (design item 3)

    private readonly record struct VisitedResult(
        double HashSetMs,
        double HashSetChurnMb,
        double BitMs,
        double BitChurnMb
    );

    private static VisitedResult ProbeVisitedGuard()
    {
        // 600k enter/exit of the cycle guard on the S formula cells (the cells that recurse). String set vs a
        // per-slot bit over the same paged layout.
        var strKeys = new (string, string)[FormulaRows * 2];
        var cols = new int[FormulaRows * 2];
        var rows = new int[FormulaRows * 2];
        var kw = 0;
        for (var r = 1; r <= FormulaRows; r++)
        {
            strKeys[kw] = (FormulaSheet, "C" + r);
            cols[kw] = 3;
            rows[kw] = r;
            kw++;
            strKeys[kw] = (FormulaSheet, "D" + r);
            cols[kw] = 4;
            rows[kw] = r;
            kw++;
        }

        // Pre-build both structures OUTSIDE the measured region: the HashSet's buckets and the store's pages
        // already exist in the running engine (the visited-bit rides on the value store's pages), so we measure
        // the steady-state op cost + incremental allocation of enter/exit only — not one-time scaffolding.
        var set = new HashSet<(string, string)>();
        var visited = new PlainSheetStore(4, FormulaRows);
        for (var i = 0; i < cols.Length; i++)
        {
            visited.SetVisited(cols[i], rows[i]); // touch every page once so it is allocated up front
            visited.ClearVisited(cols[i], rows[i]);
        }

        double bestHs = double.MaxValue;
        long bestHsChurn = long.MaxValue;
        for (var t = 0; t < PureTrials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            for (var i = 0; i < strKeys.Length; i++)
            {
                set.Add(strKeys[i]);
                set.Remove(strKeys[i]);
            }

            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            bestHs = Math.Min(bestHs, sw.Elapsed.TotalMilliseconds);
            bestHsChurn = Math.Min(bestHsChurn, churn);
            GC.KeepAlive(set);
        }

        double bestBit = double.MaxValue;
        long bestBitChurn = long.MaxValue;
        for (var t = 0; t < PureTrials; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            for (var i = 0; i < cols.Length; i++)
            {
                visited.SetVisited(cols[i], rows[i]);
                visited.ClearVisited(cols[i], rows[i]);
            }

            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            bestBit = Math.Min(bestBit, sw.Elapsed.TotalMilliseconds);
            bestBitChurn = Math.Min(bestBitChurn, churn);
            GC.KeepAlive(visited);
        }

        return new VisitedResult(bestHs, bestHsChurn / 1048576d, bestBit, bestBitChurn / 1048576d);
    }

    // ============================================================ sparsity (design item 4)

    private readonly record struct SparseResult(
        string Label,
        int TouchedPages,
        double PageDataMb,
        double PointerMb,
        double DenseTotalMb,
        double BaselineDictMb
    );

    private static List<SparseResult> ProbeSparsity()
    {
        const int cells = 10_000;
        const int maxCol = 702; // A..ZZ
        const int maxRow = 1_000_000;

        var results = new List<SparseResult>();

        // Distribution 1: uniform-random scatter (near worst case for page granularity).
        results.Add(RunSparse("scatter uniform", GenScatter(cells, maxCol, maxRow, seed: 12345)));

        // Distribution 2: clustered — 100 blocks of 100 contiguous rows in random columns (realistic table).
        results.Add(RunSparse("clustered blocks", GenClustered(cells, maxCol, maxRow, seed: 6789)));

        // Distribution 3: a few dense columns (whole-column style) — 10 columns x 1000 contiguous rows.
        results.Add(RunSparse("dense columns", GenDenseColumns(cells, maxRow: 1000)));

        return results;
    }

    private static SparseResult RunSparse(string label, (int col, int row)[] cellsAt)
    {
        var maxCol = 0;
        var maxRow = 0;
        foreach (var (c, r) in cellsAt)
        {
            if (c > maxCol)
                maxCol = c;
            if (r > maxRow)
                maxRow = r;
        }

        // Dense paged store — byte-accurate analytic footprint (see Footprint()).
        var store = new PlainSheetStore(maxCol, maxRow);
        foreach (var (c, r) in cellsAt)
        {
            store.Set(c, r, ComputedValue.Number(c * 1e6 + r));
        }

        var (pages, pageBytes, pointerBytes) = store.Footprint();
        GC.KeepAlive(store);

        // Baseline dict retained — measured (no simple analytic; heap delta is the honest proxy here).
        var baselineMb = MeasureRetained(() =>
        {
            var dict = new ConcurrentDictionary<(string, string), ComputedValue>();
            foreach (var (c, r) in cellsAt)
            {
                dict[(FormulaSheet, c + "_" + r)] = ComputedValue.Number(c * 1e6 + r);
            }

            return dict;
        });

        var pageMb = pageBytes / 1048576d;
        var pointerMb = pointerBytes / 1048576d;
        return new SparseResult(label, pages, pageMb, pointerMb, pageMb + pointerMb, baselineMb);
    }

    private static (int, int)[] GenScatter(int n, int maxCol, int maxRow, int seed)
    {
        var rnd = new Random(seed);
        var arr = new (int, int)[n];
        for (var i = 0; i < n; i++)
        {
            arr[i] = (rnd.Next(1, maxCol + 1), rnd.Next(1, maxRow + 1));
        }

        return arr;
    }

    private static (int, int)[] GenClustered(int n, int maxCol, int maxRow, int seed)
    {
        var rnd = new Random(seed);
        var arr = new (int, int)[n];
        const int blockRows = 100;
        var blocks = n / blockRows;
        var w = 0;
        for (var b = 0; b < blocks; b++)
        {
            var col = rnd.Next(1, maxCol + 1);
            var startRow = rnd.Next(1, maxRow - blockRows + 1);
            for (var k = 0; k < blockRows && w < n; k++)
            {
                arr[w++] = (col, startRow + k);
            }
        }

        while (w < n)
        {
            arr[w++] = (1, w);
        }

        return arr;
    }

    private static (int, int)[] GenDenseColumns(int n, int maxRow)
    {
        var arr = new (int, int)[n];
        var perCol = maxRow;
        var cols = n / perCol;
        var w = 0;
        for (var c = 1; c <= cols && w < n; c++)
        {
            for (var r = 1; r <= perCol && w < n; r++)
            {
                arr[w++] = (c, r);
            }
        }

        while (w < n)
        {
            arr[w++] = (cols + 1, w);
        }

        return arr;
    }

    // ============================================================ equivalence sanity

    private static void AssertEquivalence(AccessSequence seq, params Measured[] variants)
    {
        foreach (var v in variants)
        {
            if (Math.Abs(v.Sum - seq.ExpectedSum) > 1e-6)
            {
                throw new InvalidOperationException(
                    $"Dense-store spike DIVERGED: variant sum {v.Sum:R} != expected {seq.ExpectedSum:R} "
                        + $"(delta {v.Sum - seq.ExpectedSum:R}). Addressing is wrong; the numbers would lie."
                );
            }
        }

        Console.WriteLine(
            $"Equivalence sanity: all {variants.Length} variants sum to {seq.ExpectedSum:R} over 1.0M accesses. OK."
        );
        Console.WriteLine();
    }

    // ============================================================ retained-memory helper

    private static double MeasureRetained(Func<object> build)
    {
        GcQuiesce();
        var before = GC.GetTotalMemory(forceFullCollection: true);
        var obj = build();
        // Collect transient build garbage BEFORE sampling, so the delta is the retained live set only.
        GcQuiesce();
        var after = GC.GetTotalMemory(forceFullCollection: true);
        GC.KeepAlive(obj);
        return Math.Max(0, (after - before) / 1048576d);
    }

    private static void GcQuiesce()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    // ============================================================ store implementations

    // ---- variant (b): plain single-threaded paged store ----
    private sealed class PlainStore(PlainSheetStore s, PlainSheetStore data)
    {
        public PlainSheetStore Sheet(int handle) => handle == HandleS ? s : data;
    }

    private sealed class PlainSheetStore
    {
        // Page table: _cols[col] is a page-pointer array indexed by pageIndex (row-1)/PageRows. Outer sized to
        // maxCol (small); inner lazily created, sized to the column's max page from the row bound.
        private readonly PlainPage?[]?[] _cols;
        private readonly int _maxPageIndex;

        public PlainSheetStore(int maxCol, int maxRow)
        {
            _cols = new PlainPage?[maxCol + 1][];
            _maxPageIndex = maxRow >> PageShift;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private PlainPage PageFor(int col, int row)
        {
            var pages = _cols[col] ??= new PlainPage?[_maxPageIndex + 1];
            return pages[row >> PageShift] ??= new PlainPage();
        }

        public bool TryGet(int col, int row, out ComputedValue value)
        {
            var pages = _cols[col];
            if (pages is not null)
            {
                var page = pages[row >> PageShift];
                if (page is not null)
                {
                    var slot = row & PageMask;
                    if ((page.Present[slot >> 6] & (1UL << (slot & 63))) != 0)
                    {
                        value = page.Values[slot];
                        return true;
                    }
                }
            }

            value = default;
            return false;
        }

        public void Set(int col, int row, ComputedValue value)
        {
            var page = PageFor(col, row);
            var slot = row & PageMask;
            page.Values[slot] = value;
            page.Present[slot >> 6] |= 1UL << (slot & 63);
        }

        // Visited-bit reuse of the same paged layout (design item 3).
        public void SetVisited(int col, int row)
        {
            var page = PageFor(col, row);
            var slot = row & PageMask;
            page.Visited[slot >> 6] |= 1UL << (slot & 63);
        }

        public void ClearVisited(int col, int row)
        {
            var page = PageFor(col, row);
            var slot = row & PageMask;
            page.Visited[slot >> 6] &= ~(1UL << (slot & 63));
        }

        public int TouchedPageCount()
        {
            var n = 0;
            foreach (var pages in _cols)
            {
                if (pages is null)
                {
                    continue;
                }

                foreach (var p in pages)
                {
                    if (p is not null)
                    {
                        n++;
                    }
                }
            }

            return n;
        }

        // Analytic (byte-accurate) footprint: pages actually allocated + the page-pointer arrays. GetTotalMemory
        // over-reports these 24 KB arrays ~2x (GC reservation), so this is the trustworthy store-size number.
        public (int Pages, long PageBytes, long PointerBytes) Footprint()
        {
            long pointer = 0;
            var pages = 0;
            foreach (var inner in _cols)
            {
                if (inner is null)
                {
                    continue;
                }

                pointer += (long)inner.Length * 8; // one page-pointer array per touched column
                foreach (var p in inner)
                {
                    if (p is not null)
                    {
                        pages++;
                    }
                }
            }

            // Per page: Values (1024 * sizeof(ComputedValue)) + Present + Visited bitmaps (16 ulongs each).
            long perPage =
                PageRows * (long)Unsafe.SizeOf<ComputedValue>() + 2L * (PageRows / 64) * 8;
            return (pages, pages * perPage, pointer);
        }
    }

    private sealed class PlainPage
    {
        public readonly ComputedValue[] Values = new ComputedValue[PageRows];
        public readonly ulong[] Present = new ulong[PageRows / 64];
        public readonly ulong[] Visited = new ulong[PageRows / 64];
    }

    // ---- variant (a): concurrent striped store — seqlock per page (lock-free readers, CAS-guarded writers) ----
    private sealed class SeqStore(SeqSheetStore s, SeqSheetStore data)
    {
        public SeqSheetStore Sheet(int handle) => handle == HandleS ? s : data;
    }

    private sealed class SeqSheetStore
    {
        private readonly SeqPage?[]?[] _cols;
        private readonly int _maxPageIndex;
        private readonly object _growLock = new();

        public SeqSheetStore(int maxCol, int maxRow)
        {
            _cols = new SeqPage?[maxCol + 1][];
            _maxPageIndex = maxRow >> PageShift;
        }

        public bool TryGet(int col, int row, out ComputedValue value)
        {
            // Publication-safe read of the page pointer (Volatile), then a lock-free SEQLOCK read of the slot.
            var pages = Volatile.Read(ref _cols[col]);
            if (pages is not null)
            {
                var page = Volatile.Read(ref pages[row >> PageShift]);
                if (page is not null)
                {
                    return page.TryReadSlot(row & PageMask, out value);
                }
            }

            value = default;
            return false;
        }

        public void Set(int col, int row, ComputedValue value)
        {
            var page = GetOrAddPage(col, row);
            page.WriteSlot(row & PageMask, value);
        }

        private SeqPage GetOrAddPage(int col, int row)
        {
            var pi = row >> PageShift;
            var pages = Volatile.Read(ref _cols[col]);
            if (pages is null)
            {
                lock (_growLock)
                {
                    pages = _cols[col];
                    if (pages is null)
                    {
                        pages = new SeqPage?[_maxPageIndex + 1];
                        Volatile.Write(ref _cols[col], pages); // publish fully-formed array
                    }
                }
            }

            var page = Volatile.Read(ref pages[pi]);
            if (page is null)
            {
                lock (_growLock)
                {
                    page = pages[pi];
                    if (page is null)
                    {
                        page = new SeqPage();
                        Volatile.Write(ref pages[pi], page); // publish fully-formed page
                    }
                }
            }

            return page;
        }
    }

    private sealed class SeqPage
    {
        public readonly ComputedValue[] Values = new ComputedValue[PageRows];
        public readonly ulong[] Present = new ulong[PageRows / 64];
        private int _version; // even = stable, odd = a writer is mid-update
        private int _writeLock; // 0 = free, 1 = held (single-writer gate for the page)

        // Lock-free seqlock read: retry if a writer is (or was) active mid-read.
        public bool TryReadSlot(int slot, out ComputedValue value)
        {
            while (true)
            {
                var v1 = Volatile.Read(ref _version);
                if ((v1 & 1) != 0)
                {
                    continue; // writer in progress — spin
                }

                var present = (Volatile.Read(ref Present[slot >> 6]) & (1UL << (slot & 63))) != 0;
                var val = Values[slot];

                Interlocked.MemoryBarrier();
                var v2 = Volatile.Read(ref _version);
                if (v1 == v2)
                {
                    value = present ? val : default;
                    return present;
                }
                // version moved under us -> torn read possible, retry
            }
        }

        public void WriteSlot(int slot, ComputedValue value)
        {
            // Single-writer gate (page granularity): concurrent writers to the same page serialize here; writers
            // to different pages never contend. Writes are the 600k minority and already do expensive AST work.
            while (Interlocked.CompareExchange(ref _writeLock, 1, 0) != 0) { }

            var v = _version;
            Volatile.Write(ref _version, v + 1); // -> odd: readers will retry
            Values[slot] = value;
            Present[slot >> 6] |= 1UL << (slot & 63);
            Volatile.Write(ref _version, v + 2); // -> even: stable
            Volatile.Write(ref _writeLock, 0);
        }
    }
}
