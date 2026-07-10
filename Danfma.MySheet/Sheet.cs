using System.Collections;
using System.Diagnostics.CodeAnalysis;
using Danfma.MySheet.Expressions;
using MemoryPack;

namespace Danfma.MySheet;

[MemoryPackable]
public sealed partial class Sheet : IEnumerable<KeyValuePair<string, Expression>>
{
    public required string Name { get; init; }

    // 0-based insertion order, used by the SHEET function.
    public int Index { get; init; }

    // The cell store is a PRIVATE field carrying the serialized member, at the exact declaration position the
    // public `Cells` property held before — MemoryPack orders members by declaration, so this keeps member #3
    // in place. In memory it is now numeric-keyed (CellStore: an (int,int) dense map + a non-A1 overflow) to
    // shed the ~35MB of duplicate id strings the old Dictionary<string,Expression> retained at K1 scale, but
    // the wire schema is UNCHANGED: CellStoreFormatter serializes/deserializes the historical
    // Dictionary<string, Expression> byte-identically (proven by the pre-namespaces fixture + the golden wire
    // test). The initializer runs only for a fresh `new Sheet`; MemoryPack bypasses field initializers on
    // deserialize, but every serialized file carries this member (the formatter builds the store), so it is
    // never left null.
    [MemoryPackInclude]
    [CellStoreFormatter]
    private CellStore _cells = new();

    /// <summary>
    /// Read-only view of the sheet's populated cells. Mutation goes through the write choke point
    /// (<see cref="SetCell"/>, reached via the indexer <c>set</c>) and <see cref="Remove"/> — the two, and only,
    /// paths that change the cell store. Ids are exposed as their A1 strings (derived on demand for the
    /// numerically-keyed cells — a cold enumeration path).
    /// </summary>
    [MemoryPackIgnore]
    public IReadOnlyDictionary<string, Expression> Cells => _cells;

    // The write-maintained structural index (whole-column scale), owned per-Sheet because the write choke
    // point lives here. Runtime-only: [MemoryPackIgnore] so it never touches the wire schema, and lazily
    // created race-free via GetStructuralIndex (never `= new()` on the field — MemoryPack bypasses field
    // initializers on deserialize, so a loaded sheet starts with a null index and rebuilds it once on its
    // first open-range read). Built lazily, then kept up to date by SetCell/Remove; it SURVIVES
    // Workbook.InvalidateCache (structure is orthogonal to the value caches).
    [MemoryPackIgnore]
    private SheetStructuralIndex? _structuralIndex;

    /// <summary>
    /// The sheet's structural index, created (empty, unbuilt) on first access and reused for the sheet's life.
    /// The first read of a map on it triggers its one O(N) build; every <see cref="SetCell"/>/<see cref="Remove"/>
    /// keeps it current thereafter. Internal — part of the evaluation contract, not host API.
    /// </summary>
    internal SheetStructuralIndex GetStructuralIndex()
    {
        var existing = _structuralIndex;
        if (existing is not null)
        {
            return existing;
        }

        var created = new SheetStructuralIndex(this);
        return Interlocked.CompareExchange(ref _structuralIndex, created, null) ?? created;
    }

    /// <summary>The already-created structural index, or <c>null</c> when none has been created yet — a
    /// side-effect-free peek (never creates or builds). Test/introspection only.</summary>
    internal SheetStructuralIndex? PeekStructuralIndex() => _structuralIndex;

    // Monotonic counter of STRUCTURAL edits (a formula added/removed/changed) on this sheet, maintained by the
    // SetCell/Remove choke point. It bumps only when the dependency STRUCTURE can change — a formula on either
    // side of the edit — NOT for a pure value edit (literal↔literal). The reverse dependency graph
    // (RecalculationEngine) snapshots it per-sheet to know when its graph went stale and must be rebuilt, so a
    // FORMULA edit at runtime stays correct while value edits keep the cheap evict-and-pull path. Runtime-only
    // ([MemoryPackIgnore]): a loaded workbook starts at 0 and any engine built after Load rebuilds from the
    // current state anyway. Single-thread edit contract (same as the structural index), so a plain increment.
    [MemoryPackIgnore]
    private long _structuralVersion;

    /// <summary>The count of structural (formula-shape) edits this sheet has seen; the reverse dependency graph
    /// uses it to detect that it went stale. Internal — part of the recalculation contract, not host API.</summary>
    internal long StructuralVersion => _structuralVersion;

    [MemoryPackIgnore]
    public int Count => _cells.Count;

    /// <summary>
    /// Reserves dense-store capacity ahead of a bulk load (used by the .xlsx reader with the
    /// worksheet's dimension hint). Purely an allocation optimization — never affects content.
    /// </summary>
    internal void EnsureCellCapacity(int capacity) => _cells.EnsureDenseCapacity(capacity);

    [MemoryPackIgnore]
    public Expression this[string key]
    {
        get => _cells.TryGetValue(key, out var cell) ? cell : BlankValue.Instance;
        set => SetCell(key, value);
    }

    /// <summary>
    /// The single write path into the cell store: the public indexer <c>set</c> delegates here, so every cell
    /// insert or overwrite funnels through one method. It maintains the write-maintained structural index in
    /// place — an insert of a NEW id is recorded incrementally (an in-order append stays O(1); an out-of-order
    /// insert dirties only its bucket), while OVERWRITING an existing id leaves the index untouched (the id is
    /// already indexed). The index is only touched once it exists: before the sheet's first open-range read
    /// there is nothing to maintain (the eventual lazy build captures every current cell). It is the documented
    /// attach point for the future reverse dependency graph too. Like the indexer it replaced, it does NOT
    /// invalidate memoized values; the host calls <see cref="Workbook.InvalidateCache"/> after editing
    /// (see <see cref="Remove"/> for the symmetric delete).
    /// </summary>
    internal void SetCell(string id, Expression expr)
    {
        // Structural-change detection for the reverse dependency graph: a formula on EITHER side of the edit can
        // change the dependency structure (new formula adds edges; overwriting a formula changes/removes them),
        // so bump the version. A literal↔literal edit (the common input-value case) leaves the structure intact
        // and does NOT bump — that keeps the fast evict-and-pull path. The peek for the old value is only paid
        // when the NEW value is a literal (the else-branch below); a formula write short-circuits it, so bulk
        // formula loading stays a single store lookup.
        if (expr is not ValueExpression)
        {
            unchecked
            {
                _structuralVersion++;
            }
        }
        else if (_cells.TryGetValue(id, out var previous) && previous is not ValueExpression)
        {
            unchecked
            {
                _structuralVersion++; // formula → literal: its outgoing edges must be dropped
            }
        }

        // The store routes the id (canonical A1 → numeric dense map, else → overflow) and reports whether a
        // genuinely NEW A1 cell was inserted, so an overwrite skips index maintenance while a new cell updates
        // it. Only A1 cells carry column/row structure the index tracks (overflow ids are not addresses).
        _cells.Set(id, expr, out var addedDenseCell, out var column, out var row);

        if (addedDenseCell && _structuralIndex is { } index)
        {
            index.OnCellAdded(column, row);
        }
    }

    /// <summary>
    /// Removes a cell from the sheet, returning <c>true</c> when it existed (and <c>false</c> for a no-op).
    /// The second mutation path alongside the <see cref="SetCell"/> write choke point: it removes the id from
    /// the structural index (in place, O(bucket)) when the index exists, and — like a write — it does NOT
    /// invalidate memoized values: removing a cell changes the result of every formula that read it, so the
    /// host calls <see cref="Workbook.InvalidateCache"/> afterwards for the change to be observed, exactly as
    /// after a write. It is the delete-side attach point for the same future dependency graph as
    /// <see cref="SetCell"/>.
    /// </summary>
    public bool Remove(string id)
    {
        // Peek before removing so the reverse graph learns of a structural change: removing a FORMULA drops its
        // outgoing edges (structural → bump); removing a value cell is a value edit (the host recalculates the
        // removed address), so it leaves the structure intact and does NOT bump.
        var removedFormula =
            _cells.TryGetValue(id, out var previous) && previous is not ValueExpression;

        if (!_cells.Remove(id, out var wasDenseCell, out var column, out var row))
        {
            return false;
        }

        if (wasDenseCell && _structuralIndex is { } index)
        {
            index.OnCellRemoved(column, row);
        }

        if (removedFormula)
        {
            unchecked
            {
                _structuralVersion++;
            }
        }

        return true;
    }

    // The numeric (column,row) addresses of the A1 cells, for the structural-index build — it consumes the
    // coordinates directly instead of enumerating id strings only to re-parse them. Internal, not host API.
    [MemoryPackIgnore]
    /// <summary>
    /// The numeric addresses of this sheet's populated canonical-A1 cells, in insertion order —
    /// allocation-free enumeration (struct enumerator, no id strings). Pair with
    /// <see cref="Workbook.GetValueReader"/> for bulk extraction and <see cref="CellRef.TryFormat"/>
    /// to render ids without allocating. Non-A1 overflow ids are not listed (see
    /// <see cref="EnumerateCells"/>).
    /// </summary>
    public CellAddressCollection CellAddresses => _cells.DenseAddresses;

    /// <summary>
    /// This sheet's populated cells as <c>(Id, Column, Row)</c>, in insertion order. The canonical id
    /// is derived once per cell (one string each — for the allocation-free variant use
    /// <see cref="CellAddresses"/> + <see cref="CellRef.TryFormat"/>); non-A1 overflow ids are
    /// included with <c>Column = 0, Row = 0</c>.
    /// </summary>
    public SheetCellRefCollection EnumerateCells() => _cells.AddressedEntries;

    [MemoryPackIgnore]
    public IEnumerable<string> Keys => _cells.Keys;

    [MemoryPackIgnore]
    public IEnumerable<Expression> Values => _cells.Values;

    public IEnumerator<KeyValuePair<string, Expression>> GetEnumerator()
    {
        return _cells.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)_cells).GetEnumerator();
    }

    public bool ContainsKey(string key)
    {
        return _cells.ContainsKey(key);
    }

    public bool TryGetValue(string key, [MaybeNullWhen(false)] out Expression value)
    {
        return _cells.TryGetValue(key, out value);
    }

    /// <summary>
    /// Numeric-address read of a canonical A1 cell's stored expression — skips the id-string round trip
    /// entirely when the caller already has <c>(column, row)</c> in hand (e.g. SUBTOTAL's nested-subtotal
    /// scan over a range/open-range). Non-A1 overflow cells are not addressable this way; use the string
    /// indexer/<see cref="TryGetValue"/> for those. Internal — part of the evaluation fast path, not host API.
    /// </summary>
    internal bool TryGetCellExpressionDense(
        int column,
        int row,
        [MaybeNullWhen(false)] out Expression value
    ) => _cells.TryGetDense(column, row, out value);
}
