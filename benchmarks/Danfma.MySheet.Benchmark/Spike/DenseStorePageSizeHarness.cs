using System.Diagnostics;
using System.Runtime.CompilerServices;
using Danfma.MySheet;

namespace Danfma.MySheet.Benchmark.Spike;

// OWNER DIRECTIVE (2026-07-04): measure the row-page size tradeoff for the dense value store. A page of 1024
// rows is 1024 * sizeof(ComputedValue)(24) + presence(128B) ~ 24.6 KB; a small sheet (10 cols x 100 rows)
// allocates 10 such pages (~246 KB) for ~24 KB of useful slots. This sweep quantifies time AND store memory
// for row-page sizes 128 / 256 / 512 / 1024 over three shapes, to recommend the production default with numbers.
// The column group stays 64 (pointers only, negligible). The sparsity guard is intentionally OFF here — we are
// measuring the page-size tradeoff on non-pathological shapes.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --dense-store-pagesize
//
// Store model mirrors production (two-level columns groups[col>>6][col&63] -> column.pages[row>>pageShift] ->
// ComputedValue[pageRows] + presence bitmap), parameterized by pageShift. Single-threaded plain pages: page
// size is orthogonal to the seqlock, and the footprint/time it drives are what the owner asked to see.
public static class DenseStorePageSizeHarness
{
    private const int GroupShift = 6;
    private const int GroupSize = 1 << GroupShift; // 64
    private const int GroupMask = GroupSize - 1;
    private const int Trials = 7; // best-of-N (min)

    private static readonly int[] PageShifts = [7, 8, 9, 10]; // 128 / 256 / 512 / 1024 rows per page

    private readonly record struct Shape(string Name, (int Col, int Row)[] Cells, int Accesses);

    public static void Run()
    {
        Console.WriteLine("== Dense value store — row-page size sweep (OWNER DIRECTIVE 2026-07-04) ==");
        Console.WriteLine(
            $"Runtime {Environment.Version}. Column group fixed at {GroupSize}. "
                + $"sizeof(ComputedValue) = {Unsafe.SizeOf<ComputedValue>()} B. Best-of-{Trials} (min); "
                + "store footprint is byte-accurate analytic (pages + pointer arrays)."
        );
        Console.WriteLine();

        var shapes = new[]
        {
            BuildDense("small  10col x 100row", cols: 10, rows: 100),
            BuildDense("medium 50col x 5000row", cols: 50, rows: 5_000),
            BuildK1("K1     ~600k cells"),
        };

        foreach (var shape in shapes)
        {
            Console.WriteLine($"-- {shape.Name}  ({shape.Cells.Length:N0} cells, {shape.Accesses:N0} accesses) --");
            Console.WriteLine($"   {"page rows",-10}{"sweep ms",12}{"store KB",14}{"useful KB",12}{"waste x",10}");
            Console.WriteLine("   " + new string('-', 56));

            var usefulKb = shape.Cells.Length * (double)Unsafe.SizeOf<ComputedValue>() / 1024d;

            foreach (var pageShift in PageShifts)
            {
                var (ms, storeBytes) = Measure(shape, pageShift);
                var storeKb = storeBytes / 1024d;
                Console.WriteLine(
                    $"   {1 << pageShift,-10}{ms,12:N2}{storeKb,14:N1}{usefulKb,12:N1}{storeKb / usefulKb,10:N1}"
                );
            }

            Console.WriteLine();
        }

        Console.WriteLine("NOTE: 'store KB' = allocated pages (pageRows*24 + presence) + all pointer arrays;");
        Console.WriteLine("'useful KB' = cells * 24 B; 'waste x' = store / useful. Read it against sweep ms:");
        Console.WriteLine("smaller pages cut the small-sheet waste but add pointer/page overhead per row block.");
    }

    private static (double Ms, long StoreBytes) Measure(Shape shape, int pageShift)
    {
        var best = double.MaxValue;
        long storeBytes = 0;

        for (var t = 0; t < Trials; t++)
        {
            GcQuiesce();
            var store = new PagedStore(pageShift);
            var sw = Stopwatch.StartNew();

            // Populate every cell, then sweep-read it Accesses/Cells times (>=1) to mirror repeated reads.
            foreach (var (col, row) in shape.Cells)
            {
                store.Set(col, row, ComputedValue.Number(col * 2_000_000d + row));
            }

            double sink = 0;
            var reads = Math.Max(1, shape.Accesses / Math.Max(1, shape.Cells.Length));
            for (var pass = 0; pass < reads; pass++)
            {
                foreach (var (col, row) in shape.Cells)
                {
                    if (store.TryGet(col, row, out var value) && value.TryGetNumber(out var x))
                    {
                        sink += x;
                    }
                }
            }

            sw.Stop();
            GC.KeepAlive(sink);
            best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
            storeBytes = store.FootprintBytes();
        }

        return (best, storeBytes);
    }

    private static Shape BuildDense(string name, int cols, int rows)
    {
        var cells = new (int, int)[cols * rows];
        var w = 0;
        for (var c = 1; c <= cols; c++)
        {
            for (var r = 1; r <= rows; r++)
            {
                cells[w++] = (c, r);
            }
        }

        return new Shape(name, cells, cells.Length * 4);
    }

    // K1 population shape: Data A,B over 100k rows (cols 1,2) + S C,D over 200k rows (cols 3,4).
    private static Shape BuildK1(string name)
    {
        const int dataRows = 100_000;
        const int formulaRows = 200_000;
        var cells = new (int, int)[dataRows * 2 + formulaRows * 2];
        var w = 0;
        for (var r = 1; r <= dataRows; r++)
        {
            cells[w++] = (1, r);
            cells[w++] = (2, r);
        }

        for (var r = 1; r <= formulaRows; r++)
        {
            cells[w++] = (3, r);
            cells[w++] = (4, r);
        }

        return new Shape(name, cells, 1_000_000);
    }

    private static void GcQuiesce()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    // ---- parametric plain two-level paged store (page size = 1 << pageShift rows) ----
    private sealed class PagedStore(int pageShift)
    {
        private readonly int _pageRows = 1 << pageShift;
        private readonly int _pageMask = (1 << pageShift) - 1;
        private readonly int _pageShift = pageShift;
        private Column?[]?[] _groups = new Column?[]?[4];

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private Column? Find(int col)
        {
            var gi = col >> GroupShift;
            if (gi >= _groups.Length)
            {
                return null;
            }

            var group = _groups[gi];
            return group?[col & GroupMask];
        }

        public bool TryGet(int col, int row, out ComputedValue value)
        {
            var column = Find(col);
            if (column is not null)
            {
                return column.TryGet(row, _pageShift, _pageMask, out value);
            }

            value = default;
            return false;
        }

        public void Set(int col, int row, ComputedValue value)
        {
            var gi = col >> GroupShift;
            if (gi >= _groups.Length)
            {
                var size = _groups.Length;
                while (gi >= size)
                {
                    size <<= 1;
                }

                var grown = new Column?[]?[size];
                Array.Copy(_groups, grown, _groups.Length);
                _groups = grown;
            }

            var group = _groups[gi] ??= new Column?[GroupSize];
            var column = group[col & GroupMask] ??= new Column(_pageRows);
            column.Set(row, value, _pageShift, _pageMask);
        }

        public long FootprintBytes()
        {
            long pointer = (long)_groups.Length * 8;
            long pageBytes = 0;
            var perPage = (long)_pageRows * Unsafe.SizeOf<ComputedValue>() + _pageRows / 64 * 8;

            foreach (var group in _groups)
            {
                if (group is null)
                {
                    continue;
                }

                pointer += (long)group.Length * 8;
                foreach (var column in group)
                {
                    if (column is null)
                    {
                        continue;
                    }

                    var (pages, pagePtr) = column.Footprint();
                    pageBytes += pages * perPage;
                    pointer += pagePtr;
                }
            }

            return pageBytes + pointer;
        }
    }

    private sealed class Column
    {
        private Page?[] _pages = new Page?[2];
        private readonly int _pageRows;

        public Column(int pageRows) => _pageRows = pageRows;

        public bool TryGet(int row, int pageShift, int pageMask, out ComputedValue value)
        {
            var pi = row >> pageShift;
            if (pi < _pages.Length && _pages[pi] is { } page)
            {
                var slot = row & pageMask;
                if ((page.Present[slot >> 6] & (1UL << (slot & 63))) != 0)
                {
                    value = page.Values[slot];
                    return true;
                }
            }

            value = default;
            return false;
        }

        public void Set(int row, ComputedValue value, int pageShift, int pageMask)
        {
            var pi = row >> pageShift;
            if (pi >= _pages.Length)
            {
                var size = _pages.Length;
                while (pi >= size)
                {
                    size <<= 1;
                }

                var grown = new Page?[size];
                Array.Copy(_pages, grown, _pages.Length);
                _pages = grown;
            }

            var page = _pages[pi] ??= new Page(_pageRows);
            var slot = row & pageMask;
            page.Values[slot] = value;
            page.Present[slot >> 6] |= 1UL << (slot & 63);
        }

        public (int Pages, long PointerBytes) Footprint()
        {
            var pages = 0;
            foreach (var p in _pages)
            {
                if (p is not null)
                {
                    pages++;
                }
            }

            return (pages, (long)_pages.Length * 8);
        }
    }

    private sealed class Page
    {
        public readonly ComputedValue[] Values;
        public readonly ulong[] Present;

        public Page(int pageRows)
        {
            Values = new ComputedValue[pageRows];
            Present = new ulong[pageRows / 64];
        }
    }
}
