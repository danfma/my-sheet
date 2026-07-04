using System.Runtime.CompilerServices;
using Danfma.MySheet;

namespace Danfma.MySheet.Benchmark.Spike;

// Phase 2 efficacy gate (plans/allocation-hygiene-3.3.md): the adaptive first-page promotion should collapse the
// small-sheet page waste (a 10x100 sheet was paying a full ComputedValue[1024] per touched column) WITHOUT costing
// the dense K1 path anything. This probe drives the REAL SheetValueStore (not a mirror) and reads its exact
// backing footprint via the FootprintBytes diagnostic, comparing the adaptive default (InitialPageSlots=128)
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
                + "directory pointers). RowPageSize fixed at 1024."
        );
        Console.WriteLine();

        SmallSheet();
        Console.WriteLine();
        DenseK1();
    }

    private static void SmallSheet()
    {
        const int cols = 10;
        const int rows = 100;
        var usefulBytes = (long)cols * rows * Unsafe.SizeOf<ComputedValue>();

        Console.WriteLine($"-- small {cols}col x {rows}row ({cols * rows:N0} cells) --");
        Console.WriteLine($"   {"initial slots",-16}{"store KB",12}{"useful KB",12}{"waste x",10}");
        Console.WriteLine("   " + new string('-', 50));

        // Opt-out (born full, pre-3.3) then the adaptive default.
        foreach (var (label, initialSlots) in new[] { ("1024 (opt-out)", 1024), ("128 (default)", 128) })
        {
            var store = new SheetValueStore(
                new ValueStoreOptions { RowPageSize = 1024, InitialPageSlots = initialSlots }
            );
            var handle = store.HandleFor("Small");
            for (var c = 1; c <= cols; c++)
            {
                for (var r = 1; r <= rows; r++)
                {
                    store.SetDense(handle, c, r, ComputedValue.Number(c * 2_000_000d + r), tainted: false);
                }
            }

            var storeBytes = store.FootprintBytes(handle);
            Console.WriteLine(
                $"   {label,-16}{storeBytes / 1024d,12:N1}{usefulBytes / 1024d,12:N1}"
                    + $"{(double)storeBytes / usefulBytes,10:N1}"
            );
        }

        Console.WriteLine("   GATE: adaptive default waste x <= ~2 (was ~10x born-full).");
    }

    private static void DenseK1()
    {
        // K1 population shape: Data A,B over 100k rows + S C,D over 200k rows (~600k cells). A dense sheet should
        // promote every touched page all the way to full size, so the adaptive footprint matches the born-full
        // one — the small page costs the dense path nothing.
        const int dataRows = 100_000;
        const int formulaRows = 200_000;
        var cells = (long)dataRows * 2 + (long)formulaRows * 2;
        var usefulBytes = cells * Unsafe.SizeOf<ComputedValue>();

        Console.WriteLine($"-- K1 dense ({cells:N0} cells) --");
        Console.WriteLine($"   {"initial slots",-16}{"store KB",12}{"useful KB",12}{"waste x",10}");
        Console.WriteLine("   " + new string('-', 50));

        foreach (var (label, initialSlots) in new[] { ("1024 (opt-out)", 1024), ("128 (default)", 128) })
        {
            var store = new SheetValueStore(
                new ValueStoreOptions { RowPageSize = 1024, InitialPageSlots = initialSlots }
            );
            var data = store.HandleFor("Data");
            var s = store.HandleFor("S");
            for (var r = 1; r <= dataRows; r++)
            {
                store.SetDense(data, 1, r, ComputedValue.Number(r), tainted: false);
                store.SetDense(data, 2, r, ComputedValue.Number(r * 2d), tainted: false);
            }

            for (var r = 1; r <= formulaRows; r++)
            {
                store.SetDense(s, 3, r, ComputedValue.Number(r * 3d), tainted: false);
                store.SetDense(s, 4, r, ComputedValue.Number(r * 4d), tainted: false);
            }

            var storeBytes = store.FootprintBytes(data) + store.FootprintBytes(s);
            Console.WriteLine(
                $"   {label,-16}{storeBytes / 1024d,12:N1}{usefulBytes / 1024d,12:N1}"
                    + $"{(double)storeBytes / usefulBytes,10:N1}"
            );
        }

        Console.WriteLine("   GATE: adaptive default matches opt-out (dense sheet pays nothing for promotion).");
    }
}
