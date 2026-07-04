using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// The dense, paged replacement for the workbook's value cache (the old
/// <c>ConcurrentDictionary&lt;(string,string), ComputedValue&gt;</c>). It memoizes one
/// <see cref="ComputedValue"/> per cell, addressed NUMERICALLY — a sheet handle plus the (col, row) derived
/// on the fly from the canonical A1 id (see <see cref="CellAddress.TryGetColumnRow"/>, no-alloc). Runtime-only:
/// never serialized (the wire schema is untouched); the warm-start value block round-trips through
/// <see cref="EnumerateNonTainted"/>/<see cref="LoadEntry"/>.
///
/// <para><b>Layout (design item 1 — two levels on BOTH axes).</b> Per sheet a <see cref="SheetSlab"/> holds a
/// two-level column directory <c>groups[col &gt;&gt; groupShift][col &amp; groupMask]</c> (groups of 64 columns
/// by default, so a lone high-index gridless column — <c>AAAA</c> ≈ 475k — costs one group, not a 475k-pointer
/// flat array) → <c>column.pages[row &gt;&gt; pageShift]</c> → a <see cref="Page"/> of <c>ComputedValue[1024]</c>
/// by default plus its presence bitmap (a zeroed slot is ambiguous: "not computed" vs "computed blank", so
/// presence is explicit). A page's backing array actually STARTS smaller (<c>InitialPageSlots</c>, 128 by
/// default) while still covering the whole <c>RowPageSize</c>-row interval, and is PROMOTED — reallocated by
/// doubling up to <c>RowPageSize</c>, under the page seqlock — the first time a write lands past its current
/// physical size (a read past it is simply absent). This keeps a small sheet from paying a full
/// <c>ComputedValue[RowPageSize]</c> per touched column. The page/group sizes are configurable per workbook via
/// <see cref="ValueStoreOptions"/> (always powers of two, so shift/mask come free); the shifts/masks are
/// resolved once into <see cref="Geometry"/> and shared by reference with every slab/column. Group/page
/// directories start tiny and grow by doubling, publishing the new array via <see cref="Volatile"/>.
/// </para>
///
/// <para><b>Concurrency (design item 2, variant a — evaluation stays concurrent).</b> A page is a SEQLOCK:
/// readers are lock-free and re-read on a version change (the 24-byte multi-word <see cref="ComputedValue"/>
/// can tear); a single writer per page serializes on a CAS gate and bumps the version odd→even around the
/// store. The traffic is read-heavy (memoized cells are read far more than written), so a per-read lock would
/// reintroduce exactly the monitor cost this redesign removes.</para>
///
/// <para><b>Sparsity guard (design item 4 — Phase 0 risk).</b> A page is ~24.6 KB; cells scattered one-per-page
/// would balloon (10k cells over 1M rows ≈ 240 MB). Each slab tracks pages allocated vs present cells; once it
/// has more than <see cref="ValueStoreOptions.SparsityWarmupPages"/> pages yet fewer than
/// <see cref="ValueStoreOptions.SparsityMinCellsPerPage"/> cells per page (proven sparse), it stops allocating
/// NEW pages and routes further scattered cells to a per-slab dictionary.
/// Dense sheets never trip it; a pathological scatter degrades to the dictionary's footprint instead of the
/// balloon. No migration and no page teardown — only NEW page allocation is diverted, so the concurrent read
/// path is untouched.</para>
/// </summary>
internal sealed class SheetValueStore
{
    // Resolved geometry (page/group sizes → shift/mask, sparsity thresholds) captured once at construction from
    // the workbook's ValueStoreOptions. The sizes are validated powers of two; the shift is derived here so the
    // hot path stays pure shift/mask. Shared by reference with every SheetSlab/Column (one pointer, negligible
    // next to a 24.6 KB page), so a configured page/group size flows all the way to allocation.
    private sealed class Geometry
    {
        internal readonly int PageShift;
        internal readonly int PageRows;
        internal readonly int PageMask;
        internal readonly int InitialPageSlots;
        internal readonly int GroupShift;
        internal readonly int GroupSize;
        internal readonly int GroupMask;
        internal readonly int WarmupPages;
        internal readonly int MinCellsPerPage;

        internal Geometry(ValueStoreOptions options)
        {
            PageRows = options.RowPageSize;
            PageShift = System.Numerics.BitOperations.Log2((uint)options.RowPageSize);
            PageMask = options.RowPageSize - 1;
            InitialPageSlots = options.InitialPageSlots;
            GroupSize = options.ColumnGroupSize;
            GroupShift = System.Numerics.BitOperations.Log2((uint)options.ColumnGroupSize);
            GroupMask = options.ColumnGroupSize - 1;
            WarmupPages = options.SparsityWarmupPages;
            MinCellsPerPage = options.SparsityMinCellsPerPage;
        }
    }

    private readonly Geometry _geometry;

    /// <summary>Builds a store with the default geometry (row page 1024, column group 64).</summary>
    public SheetValueStore()
        : this(ValueStoreOptions.Default) { }

    /// <summary>Builds a store with the given, validated geometry. The sizes must be powers of two (enforced by
    /// <see cref="ValueStoreOptions.Validate"/>); the shift/mask are derived once here.</summary>
    public SheetValueStore(ValueStoreOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        _geometry = new Geometry(options);
    }

    // Sheet name -> dense handle (case-insensitive, like Workbook.Sheets — the name's owner). A handle is a
    // stable index into _slabs/_names for this store's life; assigned once per name under _dirLock.
    private readonly ConcurrentDictionary<string, int> _handles = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _dirLock = new();
    private int _handleCount;
    private SheetSlab?[] _slabs = new SheetSlab?[4];
    private string[] _names = new string[4]; // handle -> canonical name (warm-snapshot reverse lookup)

    // Cells that touched a volatile this epoch (design: keep the tainted set sparse — only volatile cells land
    // here — so a keyed dictionary beats a per-page bitmap that Recalculate would have to scan wholesale).
    private ConcurrentDictionary<(int Handle, int Col, int Row), byte>? _tainted;

    // Overflow for NON-A1 ids (a host may store/read a cell under a key that is not an A1 address; formulas
    // never can — the parser normalizes to A1). Dormant in normal use; preserves the old dictionary behavior
    // exactly for that path so nothing observable changes.
    private ConcurrentDictionary<(string Sheet, string Id), ComputedValue>? _overflow;
    private ConcurrentDictionary<(string Sheet, string Id), byte>? _overflowTainted;

    /// <summary>Resolves (assigning on first sight) the dense handle for a sheet name.</summary>
    public int HandleFor(string sheetName)
    {
        if (_handles.TryGetValue(sheetName, out var handle))
        {
            return handle;
        }

        lock (_dirLock)
        {
            if (_handles.TryGetValue(sheetName, out handle))
            {
                return handle;
            }

            handle = _handleCount;
            EnsureDirectoryCapacity(handle);
            _names[handle] = sheetName;
            _handleCount = handle + 1;
            _handles[sheetName] = handle; // publish last: a reader that has the handle can index _slabs/_names
            return handle;
        }
    }

    private void EnsureDirectoryCapacity(int handle)
    {
        if (handle < _slabs.Length)
        {
            return;
        }

        var size = _slabs.Length;
        while (handle >= size)
        {
            size <<= 1;
        }

        var slabs = new SheetSlab?[size];
        Array.Copy(_slabs, slabs, _slabs.Length);
        var names = new string[size];
        Array.Copy(_names, names, _names.Length);

        Volatile.Write(ref _names, names);
        Volatile.Write(ref _slabs, slabs); // publish slabs last (readers gate on it)
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private SheetSlab? PeekSlab(int handle)
    {
        var slabs = Volatile.Read(ref _slabs);
        return handle < slabs.Length ? Volatile.Read(ref slabs[handle]) : null;
    }

    private SheetSlab SlabFor(int handle)
    {
        var existing = PeekSlab(handle);
        if (existing is not null)
        {
            return existing;
        }

        lock (_dirLock)
        {
            var slabs = _slabs;
            var current = slabs[handle];
            if (current is not null)
            {
                return current;
            }

            var created = new SheetSlab(_geometry);
            Volatile.Write(ref slabs[handle], created);
            return created;
        }
    }

    // === Dense (A1) path ================================================================================

    /// <summary>Lock-free read of a memoized cell. <paramref name="col"/>/<paramref name="row"/> are 1-based.</summary>
    public bool TryGetDense(int handle, int col, int row, out ComputedValue value)
    {
        var slab = PeekSlab(handle);
        if (slab is not null)
        {
            return slab.TryGet(col, row, out value);
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Fast-path materialization of a fully-present closed column slice into a pre-sized, column-major
    /// destination array (the block-copy path of <see cref="RangeSnapshot.Build"/>). Returns true only when
    /// EVERY covered cell of the column is present and the per-page seqlock verified the copy — the caller then
    /// skips the per-cell reads for this column; false leaves <paramref name="dest"/> untouched for its
    /// per-cell fallback. <paramref name="destBase"/> is the destination index of (col, minRow).
    /// </summary>
    public bool TryBlockCopyColumn(
        int handle,
        int col,
        int minRow,
        int maxRow,
        ComputedValue[] dest,
        int destBase
    )
    {
        var slab = PeekSlab(handle);
        return slab is not null && slab.TryBlockCopyColumn(col, minRow, maxRow, dest, destBase);
    }

    /// <summary>Stores a memoized cell, marking it tainted when it touched a volatile this epoch.</summary>
    public void SetDense(int handle, int col, int row, ComputedValue value, bool tainted)
    {
        SlabFor(handle).Set(col, row, value);

        if (tainted)
        {
            (_tainted ??= NewTainted())[(handle, col, row)] = 0;
        }
    }

    private ConcurrentDictionary<(int, int, int), byte> NewTainted()
    {
        var created = new ConcurrentDictionary<(int, int, int), byte>();
        return Interlocked.CompareExchange(ref _tainted, created, null) ?? created;
    }

    // === Overflow (non-A1) path =========================================================================

    public bool TryGetOverflow(string sheetName, string id, out ComputedValue value)
    {
        var overflow = _overflow;
        if (overflow is not null)
        {
            return overflow.TryGetValue((sheetName, id), out value);
        }

        value = default;
        return false;
    }

    public void SetOverflow(string sheetName, string id, ComputedValue value, bool tainted)
    {
        var overflow = _overflow ??= new ConcurrentDictionary<(string, string), ComputedValue>();
        overflow[(sheetName, id)] = value;

        if (tainted)
        {
            (_overflowTainted ??= new ConcurrentDictionary<(string, string), byte>())[(sheetName, id)] = 0;
        }
    }

    // === Epoch operations ===============================================================================

    /// <summary>Drops every memoized value (InvalidateCache), keeping the store instance and its handle map.</summary>
    public void Clear()
    {
        var slabs = _slabs;
        for (var i = 0; i < slabs.Length; i++)
        {
            slabs[i]?.Clear();
        }

        _tainted?.Clear();
        _overflow?.Clear();
        _overflowTainted?.Clear();
    }

    /// <summary>Drops only the volatile-tainted cells (Recalculate) so they recompute on the next read.</summary>
    public void DropTainted()
    {
        if (_tainted is { } tainted)
        {
            foreach (var (handle, col, row) in tainted.Keys)
            {
                PeekSlab(handle)?.ClearPresent(col, row);
            }

            tainted.Clear();
        }

        if (_overflowTainted is { } overflowTainted && _overflow is { } overflow)
        {
            foreach (var key in overflowTainted.Keys)
            {
                overflow.TryRemove(key, out _);
            }

            overflowTainted.Clear();
        }
    }

    // === Warm-start snapshot ============================================================================

    /// <summary>
    /// Enumerates every present cell that is NOT volatile-tainted, as (sheet name, canonical A1 id, value), for
    /// the warm-start save block. The id is reconstructed from (col, row); the caller filters unrepresentable
    /// kinds (Reference) via the surrogate factory.
    /// </summary>
    public IEnumerable<(string SheetName, string Id, ComputedValue Value)> EnumerateNonTainted()
    {
        var slabs = _slabs;
        var names = _names;
        var tainted = _tainted;

        for (var handle = 0; handle < slabs.Length; handle++)
        {
            var slab = slabs[handle];
            if (slab is null)
            {
                continue;
            }

            var name = names[handle];
            foreach (var (col, row, value) in slab.EnumeratePresent())
            {
                if (tainted is not null && tainted.ContainsKey((handle, col, row)))
                {
                    continue;
                }

                yield return (name, new CellAddress(col, row).ToId(), value);
            }
        }

        if (_overflow is { } overflow)
        {
            var overflowTainted = _overflowTainted;
            foreach (var (key, value) in overflow)
            {
                if (overflowTainted is not null && overflowTainted.ContainsKey(key))
                {
                    continue;
                }

                yield return (key.Sheet, key.Id, value);
            }
        }
    }

    // Test/diagnostics only: the DEFAULT sparsity-guard thresholds and page size, so a directed test on a
    // default store can prove the scatter case caps pages instead of ballooning.
    internal const int DiagnosticWarmupPages = ValueStoreOptions.DefaultSparsityWarmupPages;
    internal const int DiagnosticMinCellsPerPage = ValueStoreOptions.DefaultSparsityMinCellsPerPage;
    internal const int DiagnosticPageRows = ValueStoreOptions.DefaultRowPageSize;
    internal const int DiagnosticInitialPageSlots = ValueStoreOptions.DefaultInitialPageSlots;

    // Test/diagnostics only: the geometry this store was actually built with, so a test can prove a configured
    // page/group size flowed through instead of being silently ignored.
    internal int ConfiguredPageRows => _geometry.PageRows;
    internal int ConfiguredColumnGroupSize => _geometry.GroupSize;
    internal int ConfiguredSparsityWarmupPages => _geometry.WarmupPages;
    internal int ConfiguredSparsityMinCellsPerPage => _geometry.MinCellsPerPage;
    internal int ConfiguredInitialPageSlots => _geometry.InitialPageSlots;

    internal (int Pages, int SparseCells) Diagnostics(int handle)
    {
        var slab = PeekSlab(handle);
        return slab is null ? (0, 0) : slab.Diagnostics();
    }

    // Test/probe: the current physical slot count of the page covering (col, row) — 0 when no page is allocated
    // yet — so a directed test can prove an adaptive page promoted (or stayed small).
    internal int PagePhysicalSlots(int handle, int col, int row)
    {
        var slab = PeekSlab(handle);
        return slab is null ? 0 : slab.PagePhysicalSlots(col, row);
    }

    // Test/probe: total backing bytes of a slab's dense pages (value arrays + presence bitmaps) plus its
    // directory pointer arrays, for the adaptive-first-page footprint gate.
    internal long FootprintBytes(int handle)
    {
        var slab = PeekSlab(handle);
        return slab is null ? 0 : slab.FootprintBytes();
    }

    /// <summary>Repopulates one cell from the warm-start block (cold path; derives address internally).</summary>
    public void LoadEntry(string sheetName, string id, ComputedValue value)
    {
        var handle = HandleFor(sheetName);

        if (CellAddress.TryGetColumnRow(id, out var col, out var row))
        {
            SetDense(handle, col, row, value, tainted: false);
        }
        else
        {
            SetOverflow(sheetName, id, value, tainted: false);
        }
    }

    // === Per-sheet slab =================================================================================

    private sealed class SheetSlab
    {
        private readonly Geometry _geo;
        private Column?[]?[] _groups = new Column?[]?[4];
        private readonly object _grow = new();

        internal SheetSlab(Geometry geo) => _geo = geo;

        // Sparsity accounting (design item 4). Pages is exact (mutated under _grow); cells is an approximate
        // running numerator (Interlocked on the 0->1 presence transition), used only to decide page allocation.
        private int _pages;
        private int _cells;

        // Sparse overflow for the scatter case: (col, row) -> value once the slab is proven too sparse to keep
        // allocating pages. Published via Volatile; ConcurrentDictionary gives per-entry atomicity.
        private ConcurrentDictionary<(int Col, int Row), ComputedValue>? _sparse;
        private volatile bool _sparseMode;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private Column? FindColumn(int col)
        {
            var groups = Volatile.Read(ref _groups);
            var gi = col >> _geo.GroupShift;
            if (gi >= groups.Length)
            {
                return null;
            }

            var group = Volatile.Read(ref groups[gi]);
            return group is null ? null : Volatile.Read(ref group[col & _geo.GroupMask]);
        }

        public bool TryGet(int col, int row, out ComputedValue value)
        {
            var column = FindColumn(col);
            if (column is not null && column.TryGet(row, out value))
            {
                return true;
            }

            var sparse = Volatile.Read(ref _sparse);
            if (sparse is not null && sparse.TryGetValue((col, row), out value))
            {
                return true;
            }

            value = default;
            return false;
        }

        public void Set(int col, int row, ComputedValue value)
        {
            var slot = row & _geo.PageMask;

            // Fast path: the page already exists — a lock-free seqlock write, no directory growth.
            var column = FindColumn(col);
            if (column is not null && column.TryGetPage(row, out var page))
            {
                if (page.WriteSlot(slot, value))
                {
                    Interlocked.Increment(ref _cells);
                }

                return;
            }

            // A cell already living in the sparse dictionary stays there.
            var sparse = Volatile.Read(ref _sparse);
            if (_sparseMode && sparse is not null && sparse.ContainsKey((col, row)))
            {
                sparse[(col, row)] = value;
                return;
            }

            SetSlow(col, row, slot, value);
        }

        private void SetSlow(int col, int row, int slot, ComputedValue value)
        {
            lock (_grow)
            {
                // Re-check: another thread may have created the page or the sparse entry meanwhile.
                var column = FindColumn(col);
                if (column is not null && column.TryGetPage(row, out var existing))
                {
                    if (existing.WriteSlot(slot, value))
                    {
                        Interlocked.Increment(ref _cells);
                    }

                    return;
                }

                if (!_sparseMode && ShouldStayDense())
                {
                    var page = GetOrAddColumnLocked(col).GetOrAddPageLocked(row);
                    _pages++;
                    if (page.WriteSlot(slot, value))
                    {
                        Interlocked.Increment(ref _cells);
                    }

                    return;
                }

                // Proven sparse (or already sparse): divert this new-page cell to the dictionary.
                var sparse = _sparse ??= new ConcurrentDictionary<(int, int), ComputedValue>();
                sparse[(col, row)] = value;
                _sparseMode = true;
            }
        }

        // Density verdict, evaluated under _grow just before allocating a NEW page. Stay dense until enough
        // pages exist to judge, then require the average occupancy to clear the floor.
        private bool ShouldStayDense() =>
            _pages < _geo.WarmupPages || Volatile.Read(ref _cells) >= (long)_pages * _geo.MinCellsPerPage;

        private Column GetOrAddColumnLocked(int col)
        {
            var groups = _groups;
            var gi = col >> _geo.GroupShift;
            if (gi >= groups.Length)
            {
                var size = groups.Length;
                while (gi >= size)
                {
                    size <<= 1;
                }

                var grown = new Column?[]?[size];
                Array.Copy(groups, grown, groups.Length);
                Volatile.Write(ref _groups, grown);
                groups = grown;
            }

            var group = groups[gi];
            if (group is null)
            {
                group = new Column?[_geo.GroupSize];
                Volatile.Write(ref groups[gi], group);
            }

            var ci = col & _geo.GroupMask;
            var column = group[ci];
            if (column is null)
            {
                column = new Column(_geo);
                Volatile.Write(ref group[ci], column);
            }

            return column;
        }

        public void ClearPresent(int col, int row)
        {
            var column = FindColumn(col);
            if (column is not null && column.TryGetPage(row, out var page))
            {
                page.ClearSlot(row & _geo.PageMask);
                return;
            }

            Volatile.Read(ref _sparse)?.TryRemove((col, row), out _);
        }

        public void Clear()
        {
            lock (_grow)
            {
                _groups = new Column?[]?[4];
                _pages = 0;
                Volatile.Write(ref _cells, 0);
                _sparse = null;
                _sparseMode = false;
            }
        }

        public (int Pages, int SparseCells) Diagnostics() =>
            (_pages, Volatile.Read(ref _sparse)?.Count ?? 0);

        internal int PagePhysicalSlots(int col, int row)
        {
            var column = FindColumn(col);
            return column?.PagePhysicalSlots(row) ?? 0;
        }

        internal long FootprintBytes()
        {
            var groups = Volatile.Read(ref _groups);
            long bytes = (long)groups.Length * 8; // top-level group-pointer array
            foreach (var group in groups)
            {
                if (group is null)
                {
                    continue;
                }

                bytes += (long)group.Length * 8; // per-group column-pointer array
                foreach (var column in group)
                {
                    if (column is not null)
                    {
                        bytes += column.FootprintBytes();
                    }
                }
            }

            return bytes;
        }

        /// <summary>Block-copy fill of one column's rows [minRow, maxRow] into <paramref name="dest"/> (see
        /// <see cref="Column.TryBlockCopy"/>). A column with any covered row diverted to the sparse dictionary
        /// (no page) fails the page pre-check and returns false, so the caller's per-cell fallback — which reads
        /// the sparse dictionary too — serves it.</summary>
        public bool TryBlockCopyColumn(int col, int minRow, int maxRow, ComputedValue[] dest, int destBase)
        {
            var column = FindColumn(col);
            return column is not null && column.TryBlockCopy(minRow, maxRow, dest, destBase);
        }

        public IEnumerable<(int Col, int Row, ComputedValue Value)> EnumeratePresent()
        {
            var groups = _groups;
            for (var gi = 0; gi < groups.Length; gi++)
            {
                var group = groups[gi];
                if (group is null)
                {
                    continue;
                }

                for (var ci = 0; ci < group.Length; ci++)
                {
                    var column = group[ci];
                    if (column is null)
                    {
                        continue;
                    }

                    var col = (gi << _geo.GroupShift) | ci;
                    foreach (var (row, value) in column.EnumeratePresent())
                    {
                        yield return (col, row, value);
                    }
                }
            }

            if (_sparse is { } sparse)
            {
                foreach (var (key, value) in sparse)
                {
                    yield return (key.Col, key.Row, value);
                }
            }
        }
    }

    // === Per-column page directory ======================================================================

    private sealed class Column
    {
        private readonly Geometry _geo;
        private Page?[] _pages = new Page?[2];
        private readonly object _grow = new();

        internal Column(Geometry geo) => _geo = geo;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool TryGet(int row, out ComputedValue value)
        {
            var pages = Volatile.Read(ref _pages);
            var pi = row >> _geo.PageShift;
            if (pi < pages.Length)
            {
                var page = Volatile.Read(ref pages[pi]);
                if (page is not null)
                {
                    return page.TryReadSlot(row & _geo.PageMask, out value);
                }
            }

            value = default;
            return false;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool TryGetPage(int row, [MaybeNullWhen(false)] out Page page)
        {
            var pages = Volatile.Read(ref _pages);
            var pi = row >> _geo.PageShift;
            if (pi < pages.Length)
            {
                page = Volatile.Read(ref pages[pi]);
                return page is not null;
            }

            page = null;
            return false;
        }

        // Allocates (publishing) the page for a row. Caller holds the slab's _grow lock, so page allocation
        // across a slab is serialized (matching the seqlock's single-writer-per-page contract on growth).
        public Page GetOrAddPageLocked(int row)
        {
            var pages = _pages;
            var pi = row >> _geo.PageShift;
            if (pi >= pages.Length)
            {
                var size = pages.Length;
                while (pi >= size)
                {
                    size <<= 1;
                }

                var grown = new Page?[size];
                Array.Copy(pages, grown, pages.Length);
                Volatile.Write(ref _pages, grown);
                pages = grown;
            }

            var page = pages[pi];
            if (page is null)
            {
                page = new Page(_geo.InitialPageSlots, _geo.PageRows);
                Volatile.Write(ref pages[pi], page);
            }

            return page;
        }

        // Test/diagnostics: the physical slot count of the page covering <paramref name="row"/> (0 when no page
        // has been allocated for it yet), so a test can prove an adaptive page promoted (or did not).
        internal int PagePhysicalSlots(int row)
        {
            var pages = Volatile.Read(ref _pages);
            var pi = row >> _geo.PageShift;
            if (pi < pages.Length && Volatile.Read(ref pages[pi]) is { } page)
            {
                return page.PhysicalSlots;
            }

            return 0;
        }

        // Test/diagnostics: total backing bytes of this column's allocated pages (physical value arrays +
        // presence bitmaps) plus its page-pointer array, for the adaptive-first-page footprint probe.
        internal long FootprintBytes()
        {
            var pages = Volatile.Read(ref _pages);
            long bytes = (long)pages.Length * 8; // page-pointer array
            foreach (var page in pages)
            {
                if (page is not null)
                {
                    bytes += page.BackingBytes;
                }
            }

            return bytes;
        }

        public IEnumerable<(int Row, ComputedValue Value)> EnumeratePresent()
        {
            var pages = _pages;
            for (var pi = 0; pi < pages.Length; pi++)
            {
                var page = pages[pi];
                if (page is null)
                {
                    continue;
                }

                var baseRow = pi << _geo.PageShift;
                foreach (var (slot, value) in page.EnumeratePresent())
                {
                    yield return (baseRow + slot, value);
                }
            }
        }

        /// <summary>
        /// Fast-path fill of <paramref name="dest"/> for this column's rows [minRow, maxRow], in column-major
        /// order (row index advances by one, so <paramref name="destBase"/> is the index of minRow). Succeeds
        /// ONLY when every covered page exists and every covered slot is present, so no on-demand miss is needed;
        /// each page's slice is then block-copied under the seqlock (torn writes retried). Returns false — without
        /// partially filling from a rejected page — when any covered slot is absent or a concurrent drop races the
        /// copy, so the caller falls back to the per-cell numeric accessor (which evaluates holes on demand).
        /// </summary>
        public bool TryBlockCopy(int minRow, int maxRow, ComputedValue[] dest, int destBase)
        {
            var pages = Volatile.Read(ref _pages);
            var firstPi = minRow >> _geo.PageShift;
            var lastPi = maxRow >> _geo.PageShift;

            if (lastPi >= pages.Length)
            {
                return false; // a covered page lies beyond the directory -> absent, needs on-demand fill
            }

            // Pre-pass: prove every covered page exists and is 100% present over its covered slice, BEFORE any
            // write into dest (so a rejection leaves dest untouched for the caller's per-cell fallback).
            for (var pi = firstPi; pi <= lastPi; pi++)
            {
                var page = Volatile.Read(ref pages[pi]);
                if (page is null)
                {
                    return false;
                }

                var sLo = pi == firstPi ? minRow & _geo.PageMask : 0;
                var sHi = pi == lastPi ? maxRow & _geo.PageMask : _geo.PageMask;
                if (!page.IsRangePresent(sLo, sHi))
                {
                    return false;
                }
            }

            // Copy pass: per page, a seqlock-verified block copy with per-page retry on a version change.
            for (var pi = firstPi; pi <= lastPi; pi++)
            {
                var page = Volatile.Read(ref pages[pi])!; // proven non-null above; page slots are never nulled
                var sLo = pi == firstPi ? minRow & _geo.PageMask : 0;
                var sHi = pi == lastPi ? maxRow & _geo.PageMask : _geo.PageMask;
                var baseRow = pi << _geo.PageShift;
                var destOffset = destBase + baseRow + sLo - minRow;

                if (!page.CopyPresentSlice(sLo, sHi, dest, destOffset))
                {
                    return false; // a slot was dropped mid-window -> bail; caller re-fills this column per cell
                }
            }

            return true;
        }
    }

    // === Page: ComputedValue[1024] + presence bitmap, guarded by a per-page seqlock =======================

    private sealed class Page
    {
        // The value array and its presence bitmap start SMALL (InitialPageSlots) yet the page still addresses the
        // full RowPageSize interval (slot = row & PageMask). A write to a slot beyond the current array PROMOTES
        // the page: the backing arrays are reallocated (doubling until the slot fits, capped at _maxSlots =
        // RowPageSize), copied, and republished — all inside a single seqlock write window, so a reader either
        // retries (version denounced the swap) or observes a fully consistent generation. Both arrays are read
        // through Volatile and every index is bounds-checked against the CAPTURED array length, so a racing read
        // during a promotion can never touch a stale, shorter array out of bounds. A slot beyond the current
        // array is, by construction, "absent" — it was never written — so reads there return false without
        // promoting.
        private ComputedValue[] _values;
        private ulong[] _present; // one 64-bit word per 64 slots
        private readonly int _maxSlots; // promotion ceiling = RowPageSize (a page never grows past its row span)
        private int _version;   // even = stable, odd = a writer is mid-update
        private int _writeLock; // 0 = free, 1 = held (single-writer gate)

        internal Page(int initialSlots, int maxSlots)
        {
            _maxSlots = maxSlots;
            var slots = Math.Min(initialSlots, maxSlots);
            _values = new ComputedValue[slots];
            _present = new ulong[WordCount(slots)];
        }

        // Words needed to cover <paramref name="slots"/> presence bits. Ceil division so a sub-64 initial array
        // (e.g. 16/32 slots) still gets one full word; a power-of-two slot count ≥ 64 gives exactly slots/64.
        private static int WordCount(int slots) => (slots + 63) >> 6;

        // Test/diagnostics: the current physical slot count and total backing bytes (value array + bitmap).
        internal int PhysicalSlots => Volatile.Read(ref _values).Length;
        internal long BackingBytes
        {
            get
            {
                var values = Volatile.Read(ref _values);
                var present = Volatile.Read(ref _present);
                return (long)values.Length * Unsafe.SizeOf<ComputedValue>() + (long)present.Length * 8;
            }
        }

        // Lock-free seqlock read: retry while a writer is (or was) active, so the multi-word ComputedValue is
        // never observed torn. The arrays are captured once and both indices are bounds-checked against the
        // captured lengths, so a promotion racing this read (which the version re-check will catch) can never
        // dereference out of bounds; a slot beyond the captured array is simply absent (never written yet).
        public bool TryReadSlot(int slot, out ComputedValue value)
        {
            var word = slot >> 6;
            var bit = 1UL << (slot & 63);

            while (true)
            {
                var v1 = Volatile.Read(ref _version);
                if ((v1 & 1) != 0)
                {
                    continue; // writer in progress
                }

                var values = Volatile.Read(ref _values);
                var present = Volatile.Read(ref _present);
                var isPresent =
                    slot < values.Length
                    && word < present.Length
                    && (Volatile.Read(ref present[word]) & bit) != 0;
                var read = isPresent ? values[slot] : default;

                Interlocked.MemoryBarrier();
                if (v1 == Volatile.Read(ref _version))
                {
                    value = read;
                    return isPresent;
                }
                // version moved under us -> possible torn read, retry
            }
        }

        /// <summary>Writes a slot; returns true when it was newly present (0→1 transition), for cell counting.
        /// Promotes the backing arrays first when the slot lies beyond the current physical size — the promotion
        /// happens inside the same odd→even version window as the write, so a reader never sees the array swapped
        /// without the version denouncing it.</summary>
        public bool WriteSlot(int slot, ComputedValue value)
        {
            var word = slot >> 6;
            var bit = 1UL << (slot & 63);

            while (Interlocked.CompareExchange(ref _writeLock, 1, 0) != 0)
            {
            }

            try
            {
                var v = _version;
                Volatile.Write(ref _version, v + 1); // odd: readers retry (covers the array swap AND the write)

                if (slot >= _values.Length)
                {
                    Grow(slot);
                }

                var wasPresent = (_present[word] & bit) != 0;
                _values[slot] = value;
                Volatile.Write(ref _present[word], _present[word] | bit);
                Volatile.Write(ref _version, v + 2); // even: stable
                return !wasPresent;
            }
            finally
            {
                Volatile.Write(ref _writeLock, 0);
            }
        }

        // Reallocates the backing arrays (doubling until <paramref name="slot"/> fits, capped at _maxSlots),
        // copies the live values/bitmap, and republishes via Volatile. Called ONLY under the write lock inside an
        // odd version window, so the swap is invisible to any reader that will not retry. slot < _maxSlots always
        // (slot = row & PageMask ≤ RowPageSize-1), so the cap never leaves the slot out of bounds.
        private void Grow(int slot)
        {
            var newLen = _values.Length;
            while (slot >= newLen)
            {
                newLen <<= 1;
            }

            if (newLen > _maxSlots)
            {
                newLen = _maxSlots;
            }

            var newValues = new ComputedValue[newLen];
            Array.Copy(_values, newValues, _values.Length);
            var newPresent = new ulong[WordCount(newLen)];
            Array.Copy(_present, newPresent, _present.Length);

            // Publish the wider present bitmap BEFORE the wider value array so that any reader which observes the
            // new (longer) _values also observes a _present at least as long — belt-and-suspenders alongside the
            // reader's own per-array bounds checks and the seqlock retry.
            Volatile.Write(ref _present, newPresent);
            Volatile.Write(ref _values, newValues);
        }

        /// <summary>Clears a slot's presence (drop for Recalculate) under the same seqlock protocol. A slot beyond
        /// the promoted region was never present, so there is nothing to clear (and no array to grow).</summary>
        public void ClearSlot(int slot)
        {
            var word = slot >> 6;
            var bit = 1UL << (slot & 63);

            while (Interlocked.CompareExchange(ref _writeLock, 1, 0) != 0)
            {
            }

            try
            {
                if (slot >= _values.Length)
                {
                    return; // never promoted this far -> never present
                }

                var v = _version;
                Volatile.Write(ref _version, v + 1);
                Volatile.Write(ref _present[word], _present[word] & ~bit);
                Volatile.Write(ref _version, v + 2);
            }
            finally
            {
                Volatile.Write(ref _writeLock, 0);
            }
        }

        public IEnumerable<(int Slot, ComputedValue Value)> EnumeratePresent()
        {
            var values = Volatile.Read(ref _values);
            var present = Volatile.Read(ref _present);
            for (var word = 0; word < present.Length; word++)
            {
                var bits = present[word];
                while (bits != 0)
                {
                    var bit = System.Numerics.BitOperations.TrailingZeroCount(bits);
                    var slot = (word << 6) + bit;
                    yield return (slot, values[slot]);
                    bits &= bits - 1;
                }
            }
        }

        /// <summary>Whether EVERY slot in the inclusive range [sLo, sHi] is present (the block-copy pre-check —
        /// a hole would need on-demand evaluation, so the raw copy is only valid over a proven-full slice). A
        /// slice extending beyond the current (un-promoted) array holds absent slots, so it is not fully
        /// present — the block-copy caller then falls back to the per-cell path.</summary>
        public bool IsRangePresent(int sLo, int sHi)
        {
            var present = Volatile.Read(ref _present);
            var loWord = sLo >> 6;
            var hiWord = sHi >> 6;

            if (hiWord >= present.Length)
            {
                return false; // covered slots lie beyond the promoted region -> absent
            }

            for (var word = loWord; word <= hiWord; word++)
            {
                // Mask off the bits outside [sLo, sHi] within the first and last words; whole interior words
                // must be fully set (~0UL). A word is "all covered slots present" when required == present&required.
                var lowBit = word == loWord ? sLo & 63 : 0;
                var highBit = word == hiWord ? sHi & 63 : 63;
                var required = highBit == 63 ? ~0UL << lowBit : ((1UL << (highBit + 1)) - 1) & (~0UL << lowBit);

                if ((Volatile.Read(ref present[word]) & required) != required)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Block-copies the contiguous slice [sLo, sHi] into <paramref name="dest"/> at
        /// <paramref name="destOffset"/> under the seqlock: the raw <see cref="Array.Copy"/> escapes the
        /// per-slot protection, so the version is re-checked AFTER the copy and the copy is retried on a change
        /// (odd version = a writer is mid-update). Returns false if a concurrent CLEAR removed a covered slot
        /// during the stable window (the caller then falls back to the per-cell path, which evaluates the hole).
        /// </summary>
        public bool CopyPresentSlice(int sLo, int sHi, ComputedValue[] dest, int destOffset)
        {
            var count = sHi - sLo + 1;

            while (true)
            {
                var v1 = Volatile.Read(ref _version);
                if ((v1 & 1) != 0)
                {
                    continue; // writer in progress
                }

                var values = Volatile.Read(ref _values);

                // A slice reaching beyond the current (un-promoted) array holds absent slots; and a concurrent
                // drop (Recalculate) could have unset a covered slot after the caller's pre-check. Re-verify both
                // inside the version-stable window so a torn/absent slot never reaches the snapshot, and so the
                // raw copy below can never index the captured array out of bounds.
                if (sHi >= values.Length || !IsRangePresent(sLo, sHi))
                {
                    return false;
                }

                Array.Copy(values, sLo, dest, destOffset, count);

                Interlocked.MemoryBarrier();
                if (v1 == Volatile.Read(ref _version))
                {
                    return true; // no writer ran across the copy -> the slice is internally consistent
                }
                // version moved under us -> a writer touched this page during the copy, retry
            }
        }
    }
}
