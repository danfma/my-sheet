using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// The Layer-1 structural index for a single sheet: <c>column → ids sorted by row</c> and the symmetric
/// <c>row → ids sorted by column</c>, each built lazily and independently. It answers "which populated
/// cells live in this column/row" WITHOUT scanning every key, so a whole-column formula in a big sheet no
/// longer pays an O(N) scan per read.
///
/// <para>Structure ≠ values: this index records only WHICH cells exist, so it is orthogonal to the value
/// cache. It is dropped by <see cref="Workbook.InvalidateCache"/> (cells may have been added/removed) but
/// SURVIVES <see cref="Workbook.Recalculate"/> (a volatile refresh changes values, never which cells
/// exist). Not serialized; owned per-epoch by the <see cref="Workbook"/>.</para>
///
/// <para>Thread-safety: each map is published once under its own lock with a double-checked read, so
/// concurrent readers of the same epoch build it at most once (the build counters assert exactly that in
/// tests). This is a plain internal class — never MemoryPack-deserialized — so field initializers are safe
/// here (unlike the <see cref="Workbook"/> cache fields, which must be lazily created via Interlocked).</para>
/// </summary>
internal sealed class SheetStructuralIndex
{
    private readonly Sheet _sheet;
    private readonly object _columnsLock = new();
    private readonly object _rowsLock = new();
    private Dictionary<int, List<string>>? _columns;
    private Dictionary<int, List<string>>? _rows;

    // Test-only build counters: prove the index is built exactly once per epoch and reused (never rebuilt
    // per read, and the row map is built independently of the column map).
    internal int ColumnBuildCount;
    internal int RowBuildCount;

    public SheetStructuralIndex(Sheet sheet) => _sheet = sheet;

    /// <summary>The column index: 1-based column → the ids populated in it, ascending by row.</summary>
    public Dictionary<int, List<string>> Columns
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

                var built = Build(byColumn: true);
                ColumnBuildCount++;
                return _columns = built;
            }
        }
    }

    /// <summary>The symmetric row index: 1-based row → the ids populated in it, ascending by column.</summary>
    public Dictionary<int, List<string>> Rows
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

                var built = Build(byColumn: false);
                RowBuildCount++;
                return _rows = built;
            }
        }
    }

    // One O(N) pass over the sheet keys, bucketed by the primary axis (column or row) with the secondary
    // axis captured only to sort each bucket. The stored lists hold ids alone (memory-frugal); the row/
    // column of an id is re-derived with the no-alloc extractor on the rare paths that still need it.
    private Dictionary<int, List<string>> Build(bool byColumn)
    {
        var buckets = new Dictionary<int, List<(int Secondary, string Id)>>();

        foreach (var id in _sheet.Keys)
        {
            if (!CellAddress.TryGetColumnRow(id, out var column, out var row))
            {
                continue;
            }

            var primary = byColumn ? column : row;
            var secondary = byColumn ? row : column;

            if (!buckets.TryGetValue(primary, out var list))
            {
                buckets[primary] = list = [];
            }

            list.Add((secondary, id));
        }

        var result = new Dictionary<int, List<string>>(buckets.Count);

        foreach (var (key, list) in buckets)
        {
            list.Sort(static (a, b) => a.Secondary.CompareTo(b.Secondary));

            var ids = new List<string>(list.Count);
            foreach (var (_, id) in list)
            {
                ids.Add(id);
            }

            result[key] = ids;
        }

        return result;
    }
}
