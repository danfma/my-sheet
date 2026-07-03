using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// The admission policy for the Layer-1 <see cref="SheetStructuralIndex"/>. Production runs
/// <see cref="Admission"/>; the other two are test/benchmark levers for the whole-column harness.
/// </summary>
internal enum StructuralIndexMode
{
    /// <summary>Build the index on a sheet's SECOND open-range read this epoch; NaiveScan the first.</summary>
    Admission,

    /// <summary>Always NaiveScan — the pre-index (2.6.1) baseline path.</summary>
    ForceNaive,

    /// <summary>Build (or reuse) the index on every read — the pre-Phase-5 tree, for regression repro.</summary>
    ForceBuild,
}

/// <summary>
/// The Layer-1 structural index for a single sheet: <c>column → ids</c> and the symmetric <c>row → ids</c>,
/// each bucketized lazily and independently in one O(N) pass. It answers "which populated cells live in this
/// column/row" WITHOUT scanning every key, so a whole-column formula in a big sheet no longer pays an O(N)
/// scan per read once the index is admitted.
///
/// <para><b>Lazy per-bucket sort (Phase 5).</b> Bucketizing does NOT sort: each column's ids are stored in
/// key-enumeration order and sorted by row only on the FIRST access to that column (symmetric for rows, by
/// column). A read that touches one narrow column of a wide sheet therefore sorts only that one list, not
/// every column of the sheet. The sort is a decorate-sort (re-derives the secondary axis once per id) mutating
/// the stored list in place under the bucket lock; every access after the first sees it already sorted.</para>
///
/// <para>Structure ≠ values: this index records only WHICH cells exist, so it is orthogonal to the value
/// cache. It is dropped by <see cref="Workbook.InvalidateCache"/> (cells may have been added/removed) but
/// SURVIVES <see cref="Workbook.Recalculate"/> (a volatile refresh changes values, never which cells
/// exist). Not serialized; owned per-epoch by the <see cref="Workbook"/>.</para>
///
/// <para>Thread-safety: each map is published once under its own lock with a double-checked read, and each
/// bucket's one-time sort runs under the same lock, so concurrent readers of the same epoch build and sort
/// each list at most once (the build counters assert exactly that in tests). This is a plain internal class —
/// never MemoryPack-deserialized — so field initializers are safe here (unlike the <see cref="Workbook"/>
/// cache fields, which must be lazily created via Interlocked).</para>
/// </summary>
internal sealed class SheetStructuralIndex
{
    private static readonly List<string> Empty = [];

    private readonly Sheet _sheet;
    private readonly object _columnsLock = new();
    private readonly object _rowsLock = new();
    private Dictionary<int, List<string>>? _columns;
    private Dictionary<int, List<string>>? _rows;
    private HashSet<int>? _sortedColumns;
    private HashSet<int>? _sortedRows;

    // Test-only build counters: prove the index is bucketized exactly once per epoch and reused (never
    // rebuilt per read, and the row map is built independently of the column map).
    internal int ColumnBuildCount;
    internal int RowBuildCount;

    public SheetStructuralIndex(Sheet sheet) => _sheet = sheet;

    // The column buckets (1-based column → ids in key-enumeration order), bucketized lazily under the lock.
    private Dictionary<int, List<string>> ColumnBuckets
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

    // The symmetric row buckets (1-based row → ids), bucketized lazily and independently of the column map.
    private Dictionary<int, List<string>> RowBuckets
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

    /// <summary>The row-sorted ids of a column (sorted lazily on first access), or <c>false</c> when the
    /// column is empty.</summary>
    public bool TryGetColumn(int column, out List<string> ids)
    {
        var buckets = ColumnBuckets;

        if (!buckets.TryGetValue(column, out var list))
        {
            ids = Empty;
            return false;
        }

        SortBucket(list, _columnsLock, _sortedColumns!, column, byColumn: true);
        ids = list;
        return true;
    }

    /// <summary>The column-sorted ids of a row (sorted lazily on first access), or <c>false</c> when the row
    /// is empty.</summary>
    public bool TryGetRow(int row, out List<string> ids)
    {
        var buckets = RowBuckets;

        if (!buckets.TryGetValue(row, out var list))
        {
            ids = Empty;
            return false;
        }

        SortBucket(list, _rowsLock, _sortedRows!, row, byColumn: false);
        ids = list;
        return true;
    }

    /// <summary>The populated-cell count of a column WITHOUT forcing its sort — used only by the cache-size
    /// estimate.</summary>
    public int ColumnLength(int column) =>
        ColumnBuckets.TryGetValue(column, out var list) ? list.Count : 0;

    /// <summary>The populated-cell count of a row WITHOUT forcing its sort.</summary>
    public int RowLength(int row) => RowBuckets.TryGetValue(row, out var list) ? list.Count : 0;

    // Test hooks: prove the lazy per-bucket sort (a column/row that was never read stays unsorted).
    internal bool IsColumnSorted(int column) => _sortedColumns is { } sorted && sorted.Contains(column);

    internal bool IsRowSorted(int row) => _sortedRows is { } sorted && sorted.Contains(row);

    // One O(N) pass over the sheet keys, bucketed by the primary axis (column or row). No sort here — each
    // bucket is ordered on first access (see SortBucket). The stored lists hold ids alone (memory-frugal).
    private Dictionary<int, List<string>> Bucketize(bool byColumn)
    {
        var buckets = new Dictionary<int, List<string>>();

        foreach (var id in _sheet.Keys)
        {
            if (!CellAddress.TryGetColumnRow(id, out var column, out var row))
            {
                continue;
            }

            var primary = byColumn ? column : row;

            if (!buckets.TryGetValue(primary, out var list))
            {
                buckets[primary] = list = [];
            }

            list.Add(id);
        }

        return buckets;
    }

    // Sorts one bucket in place the first time it is accessed, by its SECONDARY axis (a column by row, a row
    // by column). Decorate-sort: the secondary axis is re-derived once per id into a scratch array so the
    // comparison never re-parses. Idempotent and guarded by the bucket lock, so concurrent first-accessors of
    // the same bucket serialize and only one sorts.
    private static void SortBucket(List<string> list, object gate, HashSet<int> sorted, int key, bool byColumn)
    {
        lock (gate)
        {
            if (!sorted.Add(key))
            {
                return;
            }

            var count = list.Count;
            var keyed = new (int Secondary, string Id)[count];

            for (var i = 0; i < count; i++)
            {
                CellAddress.TryGetColumnRow(list[i], out var column, out var row);
                keyed[i] = (byColumn ? row : column, list[i]);
            }

            Array.Sort(keyed, static (a, b) => a.Secondary.CompareTo(b.Secondary));

            for (var i = 0; i < count; i++)
            {
                list[i] = keyed[i].Id;
            }
        }
    }
}
