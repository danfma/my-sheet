using System.Runtime.CompilerServices;
using Danfma.MySheet;

namespace Danfma.MySheet.Benchmark.Spike;

// Phase 2 efficacy gate (plans/allocation-hygiene-3.3.md): the adaptive first-page promotion should collapse the
// small-sheet page waste (a 10x100 sheet was paying a full ComputedValue[1024] per touched column) WITHOUT costing
// the dense K1 path anything — neither retained footprint NOR promotion-reallocation churn (orchestrator gate:
// populate churn over born-full <= ~1.5 MB on the K1 shape). Policy: only a column's FIRST page (index 0) is born
// small (InitialPageSlots); a column that overflows into page 1 has proven dense, so later pages are born full.
// This probe drives the REAL SheetValueStore (not a mirror): retained footprint via the FootprintBytes diagnostic,
// churn via GC.GetTotalAllocatedBytes around the populate, comparing the adaptive default (InitialPageSlots=128)
// against the opt-out (InitialPageSlots == RowPageSize, i.e. the pre-3.3 born-full behavior).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --adaptive-first-page
public static class AdaptiveFirstPageHarness
{
    public static void Run()
    {
        Console.WriteLine("== Adaptive first-page promotion — efficacy gate (Phase 2) ==");
        Console.WriteLine(
            $"Runtime {Environment.Version}. sizeof(ComputedValue) = {Unsafe.SizeOf<ComputedValue>()} B. "
                + "Store footprint is the REAL SheetValueStore.FootprintBytes (value arrays + presence bitmaps + "
                + "directory pointers); churn is GC.GetTotalAllocatedBytes(precise) around the populate. "
                + "RowPageSize fixed at 1024."
        );
        Console.WriteLine();

        Report(
            "small 10col x 100row",
            cells: 10 * 100,
            "GATE: adaptive default waste x <= ~2 (was ~10x born-full).",
            store => Populate(store, "Small", firstCol: 1, cols: 10, rows: 100)
        );
        Console.WriteLine();

        Report(
            "medium 50col x 5000row",
            cells: 50 * 5_000,
            "INFO: last partial page per column may idle above the all-adaptive floor (first-page-only policy).",
            store => Populate(store, "Medium", firstCol: 1, cols: 50, rows: 5_000)
        );
        Console.WriteLine();

        Report(
            "K1 dense (~600k cells)",
            cells: 100_000 * 2 + 200_000 * 2,
            "GATE: adaptive default matches opt-out footprint AND churn delta <= ~1.5 MB (dense path pays ~nothing).",
            store =>
            {
                // K1 population shape: Data A,B over 100k rows + S C,D over 200k rows.
                Populate(store, "Data", firstCol: 1, cols: 2, rows: 100_000);
                Populate(store, "S", firstCol: 3, cols: 2, rows: 200_000);
            }
        );
    }

    private static void Populate(
        SheetValueStore store,
        string sheet,
        int firstCol,
        int cols,
        int rows
    )
    {
        var handle = store.HandleFor(sheet);
        for (var c = firstCol; c < firstCol + cols; c++)
        {
            for (var r = 1; r <= rows; r++)
            {
                store.SetDense(
                    handle,
                    c,
                    r,
                    ComputedValue.Number(c * 2_000_000d + r),
                    tainted: false
                );
            }
        }
    }

    private static void Report(
        string title,
        long cells,
        string gate,
        Action<SheetValueStore> populate
    )
    {
        var usefulBytes = cells * Unsafe.SizeOf<ComputedValue>();

        Console.WriteLine($"-- {title} ({cells:N0} cells) --");
        Console.WriteLine(
            $"   {"initial slots", -16}{"store KB", 12}{"useful KB", 12}{"waste x", 10}{"churn MB", 12}"
        );
        Console.WriteLine("   " + new string('-', 62));

        foreach (
            var (label, initialSlots) in new[] { ("1024 (opt-out)", 1024), ("128 (default)", 128) }
        )
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var store = new SheetValueStore(
                new ValueStoreOptions { RowPageSize = 1024, InitialPageSlots = initialSlots }
            );

            var before = GC.GetTotalAllocatedBytes(precise: true);
            populate(store);
            var churn = GC.GetTotalAllocatedBytes(precise: true) - before;

            long storeBytes = 0;
            foreach (var sheet in new[] { "Small", "Medium", "Data", "S" })
            {
                storeBytes += store.FootprintBytes(store.HandleFor(sheet));
            }

            Console.WriteLine(
                $"   {label, -16}{storeBytes / 1024d, 12:N1}{usefulBytes / 1024d, 12:N1}"
                    + $"{(double)storeBytes / usefulBytes, 10:N1}{churn / 1024d / 1024d, 12:N2}"
            );
        }

        Console.WriteLine($"   {gate}");
    }
}
