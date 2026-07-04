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
/// two-level column directory <c>groups[col &gt;&gt; 6][col &amp; 63]</c> (groups of 64 columns, so a lone
/// high-index gridless column — <c>AAAA</c> ≈ 475k — costs one group, not a 475k-pointer flat array) →
/// <c>column.pages[row &gt;&gt; 10]</c> → a <see cref="Page"/> of <c>ComputedValue[1024]</c> plus a 128-byte
/// presence bitmap (a zeroed slot is ambiguous: "not computed" vs "computed blank", so presence is explicit).
/// Group/page directories start tiny and grow by doubling, publishing the new array via <see cref="Volatile"/>.
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
/// has more than <see cref="WarmupPages"/> pages yet fewer than <see cref="MinCellsPerPage"/> cells per page
/// (proven sparse), it stops allocating NEW pages and routes further scattered cells to a per-slab dictionary.
/// Dense sheets never trip it; a pathological scatter degrades to the dictionary's footprint instead of the
/// balloon. No migration and no page teardown — only NEW page allocation is diverted, so the concurrent read
/// path is untouched.</para>
/// </summary>
internal sealed class SheetValueStore
{
    private const int PageShift = 10;
    private const int PageRows = 1 << PageShift; // 1024
    private const int PageMask = PageRows - 1;
    private const int PresenceWords = PageRows / 64; // 16 ulongs = 128 bytes

    private const int GroupShift = 6;
    private const int GroupSize = 1 << GroupShift; // 64 columns per group
    private const int GroupMask = GroupSize - 1;

    // Sparsity guard thresholds (design item 4). WarmupPages: don't judge density until a slab is big enough
    // that the average is meaningful. MinCellsPerPage: below this average occupancy the slab is "sparse" and
    // new pages are diverted to the dictionary. A dense column holds ~1000 cells/page; a clustered block ~100;
    // uniform scatter ~1 — so the guard fires only for the scatter shape, exactly the ballooning case.
    private const int WarmupPages = 64;
    private const int MinCellsPerPage = 4;

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

            var created = new SheetSlab();
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

    // Test/diagnostics only: the sparsity-guard thresholds and the per-sheet (dense pages, diverted cells)
    // counts, so a directed test can prove the scatter case caps pages instead of ballooning.
    internal const int DiagnosticWarmupPages = WarmupPages;
    internal const int DiagnosticMinCellsPerPage = MinCellsPerPage;
    internal const int DiagnosticPageRows = PageRows;

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
        private Column?[]?[] _groups = new Column?[]?[4];
        private readonly object _grow = new();

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
            var gi = col >> GroupShift;
            if (gi >= groups.Length)
            {
                return null;
            }

            var group = Volatile.Read(ref groups[gi]);
            return group is null ? null : Volatile.Read(ref group[col & GroupMask]);
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
            var slot = row & PageMask;

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
            _pages < WarmupPages || Volatile.Read(ref _cells) >= (long)_pages * MinCellsPerPage;

        private Column GetOrAddColumnLocked(int col)
        {
            var groups = _groups;
            var gi = col >> GroupShift;
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
                group = new Column?[GroupSize];
                Volatile.Write(ref groups[gi], group);
            }

            var ci = col & GroupMask;
            var column = group[ci];
            if (column is null)
            {
                column = new Column();
                Volatile.Write(ref group[ci], column);
            }

            return column;
        }

        public void ClearPresent(int col, int row)
        {
            var column = FindColumn(col);
            if (column is not null && column.TryGetPage(row, out var page))
            {
                page.ClearSlot(row & PageMask);
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

                    var col = (gi << GroupShift) | ci;
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
        private Page?[] _pages = new Page?[2];
        private readonly object _grow = new();

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool TryGet(int row, out ComputedValue value)
        {
            var pages = Volatile.Read(ref _pages);
            var pi = row >> PageShift;
            if (pi < pages.Length)
            {
                var page = Volatile.Read(ref pages[pi]);
                if (page is not null)
                {
                    return page.TryReadSlot(row & PageMask, out value);
                }
            }

            value = default;
            return false;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool TryGetPage(int row, [MaybeNullWhen(false)] out Page page)
        {
            var pages = Volatile.Read(ref _pages);
            var pi = row >> PageShift;
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
            var pi = row >> PageShift;
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
                page = new Page();
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

                var baseRow = pi << PageShift;
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
        private readonly ComputedValue[] _values = new ComputedValue[PageRows];
        private readonly ulong[] _present = new ulong[PresenceWords];
        private int _version;   // even = stable, odd = a writer is mid-update
        private int _writeLock; // 0 = free, 1 = held (single-writer gate)

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
            for (var word = 0; word < PresenceWords; word++)
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
