using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

/// <summary>
/// Directed tests for the dense paged value store (plans/dense-value-store-4.0.md, Phase 1). They cover the
/// three things the full suite exercises only indirectly: the per-page SEQLOCK under concurrent read/write
/// stress (the 24-byte <see cref="ComputedValue"/> is a real torn-write subject), the cycle guard's preserved
/// semantics (including that concurrent evaluation of the SAME cell across threads is NOT a false cycle — the
/// reason the guard stays thread-local rather than a shared per-slot bit), the volatile taint / Recalculate
/// epoch model over the store, and the sparsity guard's memory ceiling.
/// </summary>
public class DenseValueStoreTests
{
    // === Seqlock: no torn reads under concurrent writes ==================================================

    [Test]
    public async Task Seqlock_ConcurrentReadsDuringWrites_NeverTearsAcrossKinds()
    {
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 3;
        const int row = 500;

        // A writer flips ONE slot between a Text and a Number — the two shapes stress all three words of the
        // struct (double, object ref, kind tag). A torn read would surface as kind=Text with a null/stale ref
        // (or kind=Number with a leftover ref), which the consistency check below rejects.
        var writes = 2_000_000;
        var done = 0;
        var torn = 0;

        var writer = Task.Run(() =>
        {
            for (var i = 0; i < writes; i++)
            {
                store.SetDense(handle, col, row, ComputedValue.Text("hello"), tainted: false);
                store.SetDense(handle, col, row, ComputedValue.Number(42d), tainted: false);
            }

            Volatile.Write(ref done, 1);
        });

        var readers = new Task[4];
        for (var t = 0; t < readers.Length; t++)
        {
            readers[t] = Task.Run(() =>
            {
                while (Volatile.Read(ref done) == 0)
                {
                    if (!store.TryGetDense(handle, col, row, out var value))
                    {
                        continue; // not yet written
                    }

                    var consistent =
                        (value.TryGetText(out var text) && text == "hello")
                        || (value.TryGetNumber(out var number) && number == 42d);

                    if (!consistent)
                    {
                        Interlocked.Exchange(ref torn, 1);
                        return;
                    }
                }
            });
        }

        await Task.WhenAll(readers.Append(writer));

        await Assert.That(torn).IsEqualTo(0);
    }

    [Test]
    public async Task ConcurrentPageAllocation_EveryCellLandsWithItsValue()
    {
        // Many threads write DISTINCT cells concurrently, forcing concurrent directory growth (column groups,
        // per-column page arrays) and page publication. Every cell must read back exactly what was written.
        var store = new SheetValueStore();
        var handle = store.HandleFor("Grid");

        const int columns = 64;   // spans multiple column groups (groups are 64 wide)
        const int rows = 4096;    // spans multiple 1024-row pages per column

        Parallel.For(1, columns + 1, col =>
        {
            for (var row = 1; row <= rows; row++)
            {
                store.SetDense(handle, col, row, ComputedValue.Number(col * 1_000_000d + row), tainted: false);
            }
        });

        var mismatches = 0;
        for (var col = 1; col <= columns; col++)
        {
            for (var row = 1; row <= rows; row++)
            {
                if (!store.TryGetDense(handle, col, row, out var value)
                    || !value.TryGetNumber(out var number)
                    || number != col * 1_000_000d + row)
                {
                    mismatches++;
                }
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    // === Adaptive first-page promotion (plans/allocation-hygiene-3.3.md, Phase 2) =======================

    [Test]
    public async Task AdaptivePage_WriteBeyondInitialArray_PromotesAndPreservesValuesAndBitmap()
    {
        // A page is born with InitialPageSlots (128) physical slots but still covers the full 1024-row interval.
        // Writing a low row leaves it small; writing a row beyond the array promotes it by reallocation, and the
        // already-present low slots (values AND presence bitmap) must survive the copy.
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 5;

        store.SetDense(handle, col, 50, ComputedValue.Number(50d), tainted: false);
        await Assert.That(store.PagePhysicalSlots(handle, col, 50)).IsEqualTo(128); // still the initial size

        // Row 300 (slot 300) is beyond the 128-slot array -> promote (128 -> 256 -> 512 to fit slot 300).
        store.SetDense(handle, col, 300, ComputedValue.Number(300d), tainted: false);
        await Assert.That(store.PagePhysicalSlots(handle, col, 300)).IsEqualTo(512);

        // The pre-promotion value survived the reallocation (its bitmap bit too).
        await Assert.That(store.TryGetDense(handle, col, 50, out var kept)).IsTrue();
        await Assert.That(kept.ToDouble()).IsEqualTo(50.0);
        await Assert.That(store.TryGetDense(handle, col, 300, out var grown)).IsTrue();
        await Assert.That(grown.ToDouble()).IsEqualTo(300.0);

        // A slot inside the promoted array but never written is still absent (presence, not size, is the truth).
        await Assert.That(store.TryGetDense(handle, col, 400, out _)).IsFalse();
    }

    [Test]
    public async Task AdaptivePage_ReadBeyondCurrentArray_IsAbsent_WithoutPromoting()
    {
        // A read of a slot beyond the current physical array is simply absent — it must NOT trigger a promotion
        // (only writes grow the page).
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 2;

        store.SetDense(handle, col, 10, ComputedValue.Number(10d), tainted: false);
        await Assert.That(store.PagePhysicalSlots(handle, col, 10)).IsEqualTo(128);

        await Assert.That(store.TryGetDense(handle, col, 900, out _)).IsFalse(); // slot 900 >> 128 -> absent
        await Assert.That(store.PagePhysicalSlots(handle, col, 10)).IsEqualTo(128); // read did not promote
    }

    [Test]
    public async Task AdaptivePromotion_UnderConcurrentReads_NeverTearsAndPreservesLowSlots()
    {
        // Readers hammer an already-present low cell while a writer walks a page's rows upward, forcing repeated
        // promotions of that page's backing arrays. The seqlock must make every array swap invisible to a reader
        // except via a version bump -> readers never tear and never lose the low cell.
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 7;

        store.SetDense(handle, col, 1, ComputedValue.Number(111d), tainted: false); // the stable low cell

        var done = 0;
        var torn = 0;

        var writer = Task.Run(() =>
        {
            // 40 sweeps of rows 2..1023 in the SAME page: each sweep re-drives promotions from 128 up to 1024.
            for (var pass = 0; pass < 40; pass++)
            {
                for (var row = 2; row <= 1023; row++)
                {
                    store.SetDense(handle, col, row, ComputedValue.Number(row), tainted: false);
                }
            }

            Volatile.Write(ref done, 1);
        });

        var readers = new Task[4];
        for (var t = 0; t < readers.Length; t++)
        {
            readers[t] = Task.Run(() =>
            {
                while (Volatile.Read(ref done) == 0)
                {
                    if (store.TryGetDense(handle, col, 1, out var value)
                        && (!value.TryGetNumber(out var n) || n != 111d))
                    {
                        Interlocked.Exchange(ref torn, 1);
                        return;
                    }
                }
            });
        }

        await Task.WhenAll(readers.Append(writer));
        await Assert.That(torn).IsEqualTo(0);

        // After all promotions the page is full-size and every row reads back its value.
        await Assert.That(store.PagePhysicalSlots(handle, col, 1)).IsEqualTo(1024);
        var mismatches = 0;
        for (var row = 2; row <= 1023; row++)
        {
            if (!store.TryGetDense(handle, col, row, out var v) || !v.TryGetNumber(out var x) || x != row)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    [Test]
    public async Task BlockCopy_OverUnpromotedPage_FallsBackWhenSliceExceedsArray_ButCopiesFullSubSlice()
    {
        // The Phase 1 block-copy must coexist with un-promoted pages. A slice reaching beyond the current small
        // array holds absent slots -> TryBlockCopyColumn returns false (per-cell fallback). A fully-present slice
        // that fits inside the small array still block-copies with correct values.
        var store = new SheetValueStore();
        var handle = store.HandleFor("S");
        const int col = 4;

        for (var row = 1; row <= 100; row++) // rows 1..100 -> slots 1..100, all inside the 128-slot initial array
        {
            store.SetDense(handle, col, row, ComputedValue.Number(row * 10d), tainted: false);
        }

        await Assert.That(store.PagePhysicalSlots(handle, col, 1)).IsEqualTo(128); // never promoted

        // (a) Slice 1..200 reaches slot 200, beyond the 128-slot array -> not fully present -> fallback (false).
        var destBig = new ComputedValue[200];
        await Assert.That(store.TryBlockCopyColumn(handle, col, 1, 200, destBig, 0)).IsFalse();

        // (b) Slice 1..100 is fully present inside the small array -> block-copies with the right values.
        var dest = new ComputedValue[100];
        await Assert.That(store.TryBlockCopyColumn(handle, col, 1, 100, dest, 0)).IsTrue();
        var mismatches = 0;
        for (var i = 0; i < 100; i++)
        {
            if (!dest[i].TryGetNumber(out var n) || n != (i + 1) * 10d)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);

        // (c) After promoting the page to cover rows up to 200, the full slice now block-copies.
        for (var row = 101; row <= 200; row++)
        {
            store.SetDense(handle, col, row, ComputedValue.Number(row * 10d), tainted: false);
        }

        await Assert.That(store.PagePhysicalSlots(handle, col, 1)).IsEqualTo(256);
        var destFull = new ComputedValue[200];
        await Assert.That(store.TryBlockCopyColumn(handle, col, 1, 200, destFull, 0)).IsTrue();
        var fullMismatches = 0;
        for (var i = 0; i < 200; i++)
        {
            if (!destFull[i].TryGetNumber(out var n) || n != (i + 1) * 10d)
            {
                fullMismatches++;
            }
        }

        await Assert.That(fullMismatches).IsEqualTo(0);
    }

    [Test]
    public async Task OptOut_InitialPageSlotsEqualsRowPageSize_PageBornFullSize_NeverPromotes()
    {
        // Setting InitialPageSlots == RowPageSize opts out of promotion: pages are born full-size (the pre-3.3
        // behavior), so a high-row write never reallocates.
        var store = new SheetValueStore(
            new ValueStoreOptions { RowPageSize = 1024, InitialPageSlots = 1024 }
        );
        var handle = store.HandleFor("S");

        store.SetDense(handle, 1, 5, ComputedValue.Number(5d), tainted: false);
        await Assert.That(store.PagePhysicalSlots(handle, 1, 5)).IsEqualTo(1024); // born full

        store.SetDense(handle, 1, 1000, ComputedValue.Number(1000d), tainted: false);
        await Assert.That(store.PagePhysicalSlots(handle, 1, 1000)).IsEqualTo(1024); // unchanged
        await Assert.That(store.TryGetDense(handle, 1, 5, out var v)).IsTrue();
        await Assert.That(v.ToDouble()).IsEqualTo(5.0);
    }

    // === Cycle guard: preserved semantics ===============================================================

    [Test]
    public async Task CircularReference_ReturnsRefError_ThroughGetCellValue()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=B1", sheet);
        sheet["B1"] = ExpressionParser.Parse("=A1", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
        await Assert.That(value.TryGetError(out var error)).IsTrue();
        await Assert.That(error).IsEqualTo(Error.Ref);
    }

    [Test]
    public async Task SameCell_EvaluatedConcurrentlyAcrossThreads_IsNotAFalseCycle()
    {
        // The cycle guard is thread-local BY DESIGN: benign concurrent (re-)evaluation of the same cell on
        // different threads must not be mistaken for a circular reference. A shared per-slot visited bit would
        // break this. A1 is a slow custom function; evaluating it from many threads at once must always yield
        // the value, never #REF!.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("=SLOW()", sheet);
        workbook.RegisterFunction("SLOW", (_, _) =>
        {
            Thread.SpinWait(2000);
            return 7d;
        });

        var refErrors = 0;
        var wrong = 0;

        Parallel.For(0, 64, _ =>
        {
            // All threads hit the uncached cell at once (SLOW spins to widen the overlap), so several evaluate
            // it concurrently — the exact benign race the thread-local guard must not flag as a cycle.
            var value = workbook.GetCellValue("Sheet1", "A1");
            if (value.TryGetError(out var error) && error == Error.Ref)
            {
                Interlocked.Increment(ref refErrors);
            }
            else if (!value.TryGetNumber(out var number) || number != 7d)
            {
                Interlocked.Increment(ref wrong);
            }
        });

        await Assert.That(refErrors).IsEqualTo(0);
        await Assert.That(wrong).IsEqualTo(0);
    }

    // === Volatile taint / Recalculate epoch model =======================================================

    [Test]
    public async Task Recalculate_DropsOnlyVolatileTaintedCells_KeepingStableCached()
    {
        var stableTicks = 0;
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // STABLE1 is a non-volatile counting function; VOL1 transitively touches NOW().
        workbook.RegisterFunction("STABLE1", (_, _) => ++stableTicks);
        sheet["S1"] = ExpressionParser.Parse("=STABLE1()", sheet);
        sheet["V1"] = ExpressionParser.Parse("=NOW()+STABLE1()", sheet);

        var early = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Local);
        workbook.TimeProvider = new FixedLocalClock(early);

        var s1 = workbook.GetCellValue("Sheet1", "S1").ToDouble();
        var v1 = workbook.GetCellValue("Sheet1", "V1").ToDouble();
        await Assert.That(s1).IsEqualTo(1.0);          // STABLE1 ran once for S1
        await Assert.That(stableTicks).IsEqualTo(2);   // and once for V1

        // New epoch + a later clock: only the volatile cell drops and recomputes; the stable one stays cached.
        var late = new DateTime(2020, 6, 15, 12, 0, 0, DateTimeKind.Local);
        workbook.TimeProvider = new FixedLocalClock(late);
        workbook.Recalculate();

        var s1Again = workbook.GetCellValue("Sheet1", "S1").ToDouble();
        await Assert.That(s1Again).IsEqualTo(1.0);
        await Assert.That(stableTicks).IsEqualTo(2); // STABLE1 did NOT re-run for S1 (still cached)

        var v1Again = workbook.GetCellValue("Sheet1", "V1").ToDouble();
        await Assert.That(v1Again).IsGreaterThan(v1); // volatile re-sampled the later clock
        await Assert.That(stableTicks).IsEqualTo(3);  // V1 recomputed → STABLE1 ran once more
    }

    [Test]
    public async Task InvalidateCache_DropsEverything_AndReEvaluates()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(10);
        sheet["B1"] = ExpressionParser.Parse("=A1", sheet);

        await Assert.That(workbook.GetCellValue("Sheet1", "B1").ToDouble()).IsEqualTo(10.0);

        sheet["A1"] = new NumberValue(20);
        workbook.InvalidateCache();

        await Assert.That(workbook.GetCellValue("Sheet1", "B1").ToDouble()).IsEqualTo(20.0);
    }

    // === Sparsity guard: bounded memory =================================================================

    [Test]
    public async Task SparseScatter_CapsPagesInsteadOfBallooning_AndStaysReadable()
    {
        // 10k cells scattered over columns A..ZZ and rows up to 1,000,000. Without the guard this is ~10k
        // distinct 24.6 KB pages ≈ 240 MB (Phase 0 risk). With it, the slab stops allocating new pages once
        // proven sparse and diverts to a dictionary, so page count is capped and every cell stays readable.
        var store = new SheetValueStore();
        var handle = store.HandleFor("Scatter");

        const int cells = 10_000;
        var random = new Random(12345);
        var expected = new Dictionary<(int, int), double>(cells);

        for (var i = 0; i < cells; i++)
        {
            var col = random.Next(1, 703);       // A..ZZ
            var row = random.Next(1, 1_000_001);  // up to 1,000,000
            var value = col * 2_000_000d + row;
            store.SetDense(handle, col, row, ComputedValue.Number(value), tainted: false);
            expected[(col, row)] = value; // last write wins, mirrors the store
        }

        var (pages, sparseCells) = store.Diagnostics(handle);

        // Pages are capped at the warmup budget; the overwhelming majority of cells landed in the sparse dict.
        await Assert.That(pages).IsLessThanOrEqualTo(SheetValueStore.DiagnosticWarmupPages);
        await Assert.That(sparseCells).IsGreaterThan(9_000);

        // Analytic footprint ceiling: dense pages (~24.6 KB each) + a small dictionary — far under 20 MB.
        var pageBytes = (long)pages
            * (SheetValueStore.DiagnosticPageRows * 24L + SheetValueStore.DiagnosticPageRows / 64 * 8L);
        await Assert.That(pageBytes).IsLessThan(20L * 1024 * 1024);

        // Every distinct cell is still served correctly (dense pages OR the sparse dict).
        var mismatches = 0;
        foreach (var ((col, row), value) in expected)
        {
            if (!store.TryGetDense(handle, col, row, out var got)
                || !got.TryGetNumber(out var number)
                || number != value)
            {
                mismatches++;
            }
        }

        await Assert.That(mismatches).IsEqualTo(0);
    }

    [Test]
    public async Task DenseColumn_StaysDense_NoSparseDiversion()
    {
        // A contiguous 5,000-row column is genuinely dense (~5 full pages, thousands of cells) — it must NOT
        // trip the sparsity guard.
        var store = new SheetValueStore();
        var handle = store.HandleFor("Dense");

        for (var row = 1; row <= 5_000; row++)
        {
            store.SetDense(handle, 1, row, ComputedValue.Number(row), tainted: false);
        }

        var (pages, sparseCells) = store.Diagnostics(handle);

        await Assert.That(sparseCells).IsEqualTo(0);          // never diverted
        await Assert.That(pages).IsLessThanOrEqualTo(6);       // ~5 pages of 1024 rows
        await Assert.That(store.TryGetDense(handle, 1, 5000, out var last)).IsTrue();
        await Assert.That(last.ToDouble()).IsEqualTo(5000.0);
    }

    // A TimeProvider pinned to a fixed local instant so the NOW() epoch is deterministic.
    private sealed class FixedLocalClock(DateTime localNow) : TimeProvider
    {
        private readonly DateTimeOffset _now = new(
            DateTime.SpecifyKind(localNow, DateTimeKind.Unspecified),
            TimeSpan.Zero
        );

        public override DateTimeOffset GetUtcNow() => _now;

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
