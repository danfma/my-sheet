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
/// presence is explicit). The page/group sizes are configurable per workbook via
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

    // Test/diagnostics only: the geometry this store was actually built with, so a test can prove a configured
    // page/group size flowed through instead of being silently ignored.
    internal int ConfiguredPageRows => _geometry.PageRows;
    internal int ConfiguredColumnGroupSize => _geometry.GroupSize;
    internal int ConfiguredSparsityWarmupPages => _geometry.WarmupPages;
    internal int ConfiguredSparsityMinCellsPerPage => _geometry.MinCellsPerPage;

    internal (int Pages, int SparseCells) Diagnostics(int handle)
    {
        var slab = PeekSlab(handle);
        return slab is null ? (0, 0) : slab.Diagnostics();
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
                page = new Page(_geo.PageRows);
                Volatile.Write(ref pages[pi], page);
            }

            return page;
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
    }

    // === Page: ComputedValue[1024] + presence bitmap, guarded by a per-page seqlock =======================

    private sealed class Page
    {
        private readonly ComputedValue[] _values;
        private readonly ulong[] _present; // one 64-bit word per 64 slots (page rows are a power of two ≥ 64)
        private int _version;   // even = stable, odd = a writer is mid-update
        private int _writeLock; // 0 = free, 1 = held (single-writer gate)

        internal Page(int pageRows)
        {
            _values = new ComputedValue[pageRows];
            _present = new ulong[pageRows / 64];
        }

        // Lock-free seqlock read: retry while a writer is (or was) active, so the multi-word ComputedValue is
        // never observed torn.
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

                var present = (Volatile.Read(ref _present[word]) & bit) != 0;
                var read = _values[slot];

                Interlocked.MemoryBarrier();
                if (v1 == Volatile.Read(ref _version))
                {
                    value = present ? read : default;
                    return present;
                }
                // version moved under us -> possible torn read, retry
            }
        }

        /// <summary>Writes a slot; returns true when it was newly present (0→1 transition), for cell counting.</summary>
        public bool WriteSlot(int slot, ComputedValue value)
        {
            var word = slot >> 6;
            var bit = 1UL << (slot & 63);

            while (Interlocked.CompareExchange(ref _writeLock, 1, 0) != 0)
            {
            }

            try
            {
                var wasPresent = (_present[word] & bit) != 0;
                var v = _version;
                Volatile.Write(ref _version, v + 1); // odd: readers retry
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

        /// <summary>Clears a slot's presence (drop for Recalculate) under the same seqlock protocol.</summary>
        public void ClearSlot(int slot)
        {
            var word = slot >> 6;
            var bit = 1UL << (slot & 63);

            while (Interlocked.CompareExchange(ref _writeLock, 1, 0) != 0)
            {
            }

            try
            {
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
            for (var word = 0; word < _present.Length; word++)
            {
                var bits = _present[word];
                while (bits != 0)
                {
                    var bit = System.Numerics.BitOperations.TrailingZeroCount(bits);
                    var slot = (word << 6) + bit;
                    yield return (slot, _values[slot]);
                    bits &= bits - 1;
                }
            }
        }
    }
}
