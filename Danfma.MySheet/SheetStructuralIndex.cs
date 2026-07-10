using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

// === The structural-layers invariant (anchor note) =====================================================
// A workbook's "structure" — which cells exist, and what depends on what — is tracked in THREE separate
// places, each with its own lifetime. Read this before touching any of them:
//
//   1. SheetStructuralIndex (this file, per-Sheet) — write-maintained by Sheet.SetCell/Remove, so it is
//      always current. LIFETIME-scoped: Workbook.InvalidateCache does NOT drop it (a value refresh never
//      changes which cells exist; only an insert/delete does, and both paths update it in place).
//   2. ReverseDependencyGraph buckets (DirtyGraph/ReverseDependencyGraph.cs) — a built SNAPSHOT, not
//      write-maintained. Staleness is version-checked, not intercepted: RecalculationEngine compares each
//      Sheet.StructuralVersion (bumped on formula add/remove/change) and Workbook.NamesVersion (bumped on
//      DefineName) against the versions it was built from, and rebuilds the WHOLE graph when either moved.
//      A pure value edit bumps neither, so it never triggers a rebuild.
//   3. SheetValueStore pages (SheetValueStore.cs) — memoized computed values, EPOCH-scoped:
//      Workbook.InvalidateCache clears them unconditionally, and every read after that repopulates lazily.
//
// Consistency rule: never assume (2) or (3) reflect a structural edit until their own trigger — a version
// bump for (2), InvalidateCache for (3) — has actually fired. See the cross-references in the other two
// files.
//
/// <summary>
/// The write-maintained structural index for a single sheet: <c>column → rows</c> and the symmetric
/// <c>row → columns</c>, each bucketized lazily and independently in one O(N) pass. It answers "which populated
/// cells live in this column/row" WITHOUT scanning every key, so a whole-column formula in a big sheet no
/// longer pays an O(N) scan per read once the index exists.
///
/// <para><b>Numeric buckets (3.3).</b> Each bucket stores the SECONDARY axis as a plain <see cref="int"/> — a
/// column keeps its populated ROWS, a row keeps its populated COLUMNS — never the id string. The id is only ever
/// re-derived on a cold path (<see cref="OpenRangeReference.PopulatedIds"/> / <see cref="OpenRangeReference.Expand"/>)
/// via <see cref="CellAddress.ToId"/>; the hot open-range value path
/// (<see cref="OpenRangeReference.ExpandComputedValues"/>) reads the <c>(column, row)</c> pair straight out of the
/// index and addresses the dense store numerically, so it never parses an id. This also shrinks the index ~10× on a
/// big sheet (a 4-byte int per cell instead of a ~38-byte "C123456" string reference plus its array slot).</para>
///
/// <para><b>Lifetime, not epoch (3.0).</b> The index is built LAZILY on the first open-range read of the
/// OWNING <see cref="Sheet"/>'s life and then kept up to date by every <see cref="Sheet.SetCell"/> and
/// <see cref="Sheet.Remove"/>. It is <b>runtime-only</b> (never serialized): a workbook loaded from disk starts
/// with no index, so the first open-range read after a <see cref="Workbook.Load(string)"/> rebuilds it once for
/// that instance's life. Crucially it is NOT dropped by <see cref="Workbook.InvalidateCache"/> — a value
/// refresh changes values, never which cells exist, and neither does clearing the value caches. The only thing
/// that changes structure is a cell insert/delete, and those are exactly the two paths that maintain it here.</para>
///
/// <para><b>Adaptive maintenance.</b> Bucketizing does NOT sort: each bucket is sorted by its secondary axis
/// (a column by row, a row by column) only on the FIRST access to that bucket, and every access after sees it
/// already sorted. A new cell whose secondary axis is beyond the bucket's current last (the typical Fill /
/// append case) is an O(1) append that keeps the bucket sorted; an out-of-order insert appends and marks ONLY
/// that bucket dirty, so the next read of that ONE bucket re-sorts it (no other bucket is touched). Overwriting
/// an existing cell never reaches here (the id is already indexed). A delete removes the secondary from the
/// affected bucket in O(bucket) — rare next to appends, and it keeps the bucket sorted, so no re-sort is needed.</para>
///
/// <para>Thread-safety: each map is published once under its own lock with a double-checked read, and each
/// bucket's one-time sort and every incremental mutation run under that same lock, so concurrent readers of a
/// built map serialize against a concurrent maintenance call. Structural maintenance is expected only during
/// the edit phase (edit → <see cref="Workbook.InvalidateCache"/> → read), never concurrently with evaluation —
/// the same contract the value cache already relies on. This is a plain internal class — never
/// MemoryPack-deserialized — so field initializers are safe here (unlike the <see cref="Workbook"/>/<see
/// cref="Sheet"/> cache fields, which must be lazily created via Interlocked).</para>
/// </summary>
internal sealed class SheetStructuralIndex
{
    private static readonly List<int> Empty = [];

    private readonly Sheet _sheet;
    private readonly object _columnsLock = new();
    private readonly object _rowsLock = new();
    private Dictionary<int, List<int>>? _columns;
    private Dictionary<int, List<int>>? _rows;
    private HashSet<int>? _sortedColumns;
    private HashSet<int>? _sortedRows;

    // Test-only counters. Build: prove each map is bucketized exactly once per LIFE (never per epoch) and the
    // two maps are built independently. Append: prove an in-order write is an O(1) append with no re-sort.
    // Sort: prove an out-of-order write dirties ONE bucket and only that bucket re-sorts on its next read.
    internal int ColumnBuildCount;
    internal int RowBuildCount;
    internal int ColumnAppendCount;
    internal int RowAppendCount;
    internal int ColumnSortCount;
    internal int RowSortCount;

    public SheetStructuralIndex(Sheet sheet) => _sheet = sheet;

    // The column buckets (1-based column → rows in key-enumeration order), bucketized lazily under the lock.
    private Dictionary<int, List<int>> ColumnBuckets
    {
        get
        {
            if (_columns is { } existing)
            {
                return existing;
            }

            lock (_columnsLock)
            {
                if (_columns is { } inner)
                {
                    return inner;
                }

                var built = Bucketize(byColumn: true);
                _sortedColumns = new HashSet<int>();
                ColumnBuildCount++;
                return _columns = built;
            }
        }
    }

    // The symmetric row buckets (1-based row → columns), bucketized lazily and independently of the column map.
    private Dictionary<int, List<int>> RowBuckets
    {
        get
        {
            if (_rows is { } existing)
            {
                return existing;
            }

            lock (_rowsLock)
            {
                if (_rows is { } inner)
                {
                    return inner;
                }

                var built = Bucketize(byColumn: false);
                _sortedRows = new HashSet<int>();
                RowBuildCount++;
                return _rows = built;
            }
        }
    }

    /// <summary>The 1-based columns that have at least one populated cell (unsorted set — for open-column-side
    /// enumeration and cheap length probes).</summary>
    public IReadOnlyCollection<int> ColumnKeys => ColumnBuckets.Keys;

    /// <summary>The 1-based rows that have at least one populated cell.</summary>
    public IReadOnlyCollection<int> RowKeys => RowBuckets.Keys;

    /// <summary>The sorted ROWS populated in a column (sorted lazily on first access), or <c>false</c> when the
    /// column is empty.</summary>
    public bool TryGetColumn(int column, out List<int> rows)
    {
        var buckets = ColumnBuckets;

        if (!buckets.TryGetValue(column, out var list))
        {
            rows = Empty;
            return false;
        }

        SortBucket(list, _columnsLock, _sortedColumns!, column, byColumn: true);
        rows = list;
        return true;
    }

    /// <summary>The sorted COLUMNS populated in a row (sorted lazily on first access), or <c>false</c> when the
    /// row is empty.</summary>
    public bool TryGetRow(int row, out List<int> columns)
    {
        var buckets = RowBuckets;

        if (!buckets.TryGetValue(row, out var list))
        {
            columns = Empty;
            return false;
        }

        SortBucket(list, _rowsLock, _sortedRows!, row, byColumn: false);
        columns = list;
        return true;
    }

    /// <summary>The populated-cell count of a column WITHOUT forcing its sort — used only by the cache-size
    /// estimate.</summary>
    public int ColumnLength(int column) =>
        ColumnBuckets.TryGetValue(column, out var list) ? list.Count : 0;

    /// <summary>The populated-cell count of a row WITHOUT forcing its sort.</summary>
    public int RowLength(int row) => RowBuckets.TryGetValue(row, out var list) ? list.Count : 0;

    // Test hooks: prove the lazy per-bucket sort (a column/row that was never read stays unsorted) and the
    // dirty-marking of an out-of-order insert (the touched bucket drops out of the sorted set until re-read).
    internal bool IsColumnSorted(int column) =>
        _sortedColumns is { } sorted && sorted.Contains(column);

    internal bool IsRowSorted(int row) => _sortedRows is { } sorted && sorted.Contains(row);

    /// <summary>
    /// Incrementally records a cell that was just ADDED to the sheet (a new id, not an overwrite). No-op for a
    /// map that has not been built yet — its eventual lazy build captures every current cell in one pass. See
    /// the class remarks for the adaptive append-vs-dirty rule.
    /// </summary>
    public void OnCellAdded(int column, int row)
    {
        if (_columns is not null)
        {
            lock (_columnsLock)
            {
                if (
                    _columns is { } columns
                    && AddToBucket(columns, _sortedColumns!, primary: column, secondary: row)
                )
                {
                    ColumnAppendCount++;
                }
            }
        }

        if (_rows is not null)
        {
            lock (_rowsLock)
            {
                if (
                    _rows is { } rows
                    && AddToBucket(rows, _sortedRows!, primary: row, secondary: column)
                )
                {
                    RowAppendCount++;
                }
            }
        }
    }

    /// <summary>
    /// Incrementally records a cell that was just REMOVED from the sheet. No-op for a map not yet built. Removes
    /// the secondary from its bucket in O(bucket), preserving the bucket's sort order (so no re-sort is needed),
    /// and drops a bucket that becomes empty so <see cref="ColumnKeys"/>/<see cref="RowKeys"/> stay accurate.
    /// </summary>
    public void OnCellRemoved(int column, int row)
    {
        if (_columns is not null)
        {
            lock (_columnsLock)
            {
                if (_columns is { } columns)
                {
                    RemoveFromBucket(columns, _sortedColumns!, primary: column, secondary: row);
                }
            }
        }

        if (_rows is not null)
        {
            lock (_rowsLock)
            {
                if (_rows is { } rows)
                {
                    RemoveFromBucket(rows, _sortedRows!, primary: row, secondary: column);
                }
            }
        }
    }

    // Adds a secondary axis to its primary bucket (the caller holds the map's lock). Returns true only for an
    // in-order O(1) append onto an already-sorted bucket (the Fill case), so the caller can count it; a brand-new
    // bucket or an unsorted bucket just appends (false), and an out-of-order insert onto a sorted bucket appends
    // and drops the bucket from the sorted set (false) so its next read re-sorts it.
    private static bool AddToBucket(
        Dictionary<int, List<int>> map,
        HashSet<int> sorted,
        int primary,
        int secondary
    )
    {
        if (!map.TryGetValue(primary, out var list))
        {
            map[primary] = [secondary];
            sorted.Add(primary); // a singleton bucket is trivially sorted
            return false;
        }

        if (!sorted.Contains(primary))
        {
            list.Add(secondary); // bucket not yet sorted → append; it sorts on next read regardless
            return false;
        }

        var lastSecondary = list[^1];

        list.Add(secondary);

        if (secondary > lastSecondary)
        {
            return true; // stayed sorted: O(1) append, no re-sort
        }

        sorted.Remove(primary); // out-of-order: mark this ONE bucket dirty
        return false;
    }

    // Removes a secondary axis from its bucket (the caller holds the map's lock). Preserves sort order; drops an
    // emptied bucket and its sorted flag.
    private static void RemoveFromBucket(
        Dictionary<int, List<int>> map,
        HashSet<int> sorted,
        int primary,
        int secondary
    )
    {
        if (!map.TryGetValue(primary, out var list))
        {
            return;
        }

        list.Remove(secondary);

        if (list.Count == 0)
        {
            map.Remove(primary);
            sorted.Remove(primary);
        }
    }

    // One O(N) pass over the sheet's A1 cell addresses, bucketed by the primary axis (column or row). No sort
    // here — each bucket is ordered on first access (see SortBucket). The stored lists hold the secondary axis
    // as an int (memory-frugal: no id string, no reference). The cell store hands the numeric (column,row)
    // coordinates directly, so the build no longer re-derives an id string per cell only to re-parse it (the
    // non-A1 overflow cells carry no address and are structurally invisible, exactly as before).
    private Dictionary<int, List<int>> Bucketize(bool byColumn)
    {
        var buckets = new Dictionary<int, List<int>>();

        foreach (var (column, row) in _sheet.CellAddresses)
        {
            var primary = byColumn ? column : row;
            var secondary = byColumn ? row : column;

            if (!buckets.TryGetValue(primary, out var list))
            {
                buckets[primary] = list = [];
            }

            list.Add(secondary);
        }

        return buckets;
    }

    // Sorts one bucket in place the first time it is accessed, by its SECONDARY axis (a column by row, a row
    // by column). The secondary axis is stored as an int, so this is a plain ascending int sort — no decorate
    // step and no re-parse. Idempotent and guarded by the bucket lock, so concurrent first-accessors of the
    // same bucket serialize and only one sorts. A dirtied bucket (dropped from the sorted set by an
    // out-of-order insert) is re-sorted here on its next read, exactly as a first sort.
    private void SortBucket(
        List<int> list,
        object gate,
        HashSet<int> sorted,
        int key,
        bool byColumn
    )
    {
        lock (gate)
        {
            if (!sorted.Add(key))
            {
                return;
            }

            list.Sort();

            if (byColumn)
            {
                ColumnSortCount++;
            }
            else
            {
                RowSortCount++;
            }
        }
    }
}
