using MemoryPack;

namespace Danfma.MySheet.Expressions;

// SheetName carries the same read-side interning as CellReference (wire byte-identical; see that file).
[MemoryPackable]
public sealed partial record RangeReference(
    string StartId,
    string EndId,
    [property: InternStringFormatter] string SheetName
) : Reference
{
    // A range has no scalar value: used outside a function that accepts ranges it is a #VALUE! error.
    public override ComputedValue Evaluate(EvaluationContext context) => ComputedValue.Error(Error.Value);

    // Backwards-compatible entry point used by tests and external callers.
    public IEnumerable<Expression> Expand(Workbook workbook) =>
        Expand(new EvaluationContext(workbook));

    /// <summary>
    /// Enumerates the stored expression of every cell in the rectangle (blank cells included as
    /// <see cref="BlankValue"/>). Reversed corners (e.g. <c>B2:A1</c>) are normalized.
    /// </summary>
    public IEnumerable<Expression> Expand(EvaluationContext context)
    {
        // Two layers guard a missing sheet. LOCAL guarantee: this enumerator never throws — a missing sheet
        // yields nothing (yield break) instead of a KeyNotFoundException. STRUCTURAL resolution: every
        // reference-consuming function first runs ReferenceGuard.MissingSheet and returns #REF! before it ever
        // enumerates, so a missing sheet surfaces as #REF! (Excel parity) rather than a silently empty range.
        if (!context.Workbook.Sheets.TryGetValue(SheetName, out var sheet))
        {
            yield break;
        }

        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);

        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        for (var column = minColumn; column <= maxColumn; column++)
        {
            for (var row = minRow; row <= maxRow; row++)
            {
                yield return sheet[new CellAddress(column, row).ToId()];
            }
        }
    }

    /// <summary>Enumerates the memoized <see cref="ComputedValue"/> of every cell in the rectangle (via the
    /// workbook cache) — the allocation-free path used by range-consuming functions. Returns a concrete
    /// value-type sequence (not <see cref="IEnumerable{T}"/>): a <c>foreach</c> binds its struct enumerator by
    /// duck typing (the <see cref="List{T}"/> pattern), so the hot fold pays no per-cell interface dispatch and
    /// the expansion allocates nothing — no iterator state machine. The handle is resolved once and each cell
    /// read is numeric; the memoization, cycle guard, volatile taint and on-demand miss evaluation are
    /// unchanged (they live behind <see cref="Workbook.GetCellValueDense"/>).</summary>
    internal RangeValueSequence ExpandComputedValues(EvaluationContext context)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);

        var minColumn = Math.Min(start.Column, end.Column);
        var maxColumn = Math.Max(start.Column, end.Column);
        var minRow = Math.Min(start.Row, end.Row);
        var maxRow = Math.Max(start.Row, end.Row);

        var workbook = context.Workbook;
        var store = workbook.DenseStore;
        var handle = store.HandleFor(SheetName);

        return new RangeValueSequence(
            workbook,
            store,
            SheetName,
            handle,
            minColumn,
            maxColumn,
            minRow,
            maxRow
        );
    }

    public int RowCount =>
        Math.Abs(CellAddress.Parse(EndId).Row - CellAddress.Parse(StartId).Row) + 1;

    public int ColumnCount =>
        Math.Abs(CellAddress.Parse(EndId).Column - CellAddress.Parse(StartId).Column) + 1;

    public int TopRow => Math.Min(CellAddress.Parse(StartId).Row, CellAddress.Parse(EndId).Row);

    public int LeftColumn =>
        Math.Min(CellAddress.Parse(StartId).Column, CellAddress.Parse(EndId).Column);

    /// <summary>The memoized <see cref="ComputedValue"/> at a 1-based (row, column) position (normalized corners).</summary>
    internal ComputedValue CellComputedValueAt(EvaluationContext context, int row, int column)
    {
        var start = CellAddress.Parse(StartId);
        var end = CellAddress.Parse(EndId);
        var absoluteColumn = Math.Min(start.Column, end.Column) + column - 1;
        var absoluteRow = Math.Min(start.Row, end.Row) + row - 1;

        // Numeric address (no ToId round trip on a hit); same GetCellValue semantics via the dense twin.
        var workbook = context.Workbook;
        return workbook.GetCellValueDense(
            workbook.ResolveDenseHandle(SheetName),
            SheetName,
            absoluteColumn,
            absoluteRow
        );
    }
}

/// <summary>
/// A no-alloc, dispatch-free view over a closed rectangle's memoized cell values, returned by
/// <see cref="RangeReference.ExpandComputedValues"/>. It is a value type carrying only the resolved bounds
/// and the once-resolved sheet handle; its <see cref="Enumerator"/> reads each cell numerically through
/// <see cref="Workbook.GetCellValueDense"/> in column-major (row-inner) order — the exact order, and the
/// exact on-demand evaluation timing, of the enumerator it replaced. A <c>foreach</c> binds
/// <see cref="GetEnumerator"/> by duck typing, so no interface is boxed and no iterator object is allocated;
/// it implements <see cref="IEnumerable{T}"/> too, only for the rare cold caller that needs the interface
/// (that path boxes, exactly like <see cref="List{T}"/>).
/// </summary>
internal readonly struct RangeValueSequence(
    Workbook workbook,
    SheetValueStore store,
    string sheetName,
    int handle,
    int minColumn,
    int maxColumn,
    int minRow,
    int maxRow
) : IEnumerable<ComputedValue>
{
    public Enumerator GetEnumerator() =>
        new(workbook, store, sheetName, handle, minColumn, maxColumn, minRow, maxRow);

    IEnumerator<ComputedValue> IEnumerable<ComputedValue>.GetEnumerator() => GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>The mutable struct cursor: column-major (row-inner) over the rectangle, computing each value
    /// during <see cref="MoveNext"/> so evaluation order matches the previous <c>yield</c> path exactly.</summary>
    public struct Enumerator : IEnumerator<ComputedValue>
    {
        private readonly Workbook _workbook;
        private readonly SheetValueStore _store;
        private readonly string _sheetName;
        private readonly int _handle;
        private readonly int _minColumn;
        private readonly int _maxColumn;
        private readonly int _minRow;
        private readonly int _maxRow;
        private int _column;
        private int _row;
        private ComputedValue _current;

        internal Enumerator(
            Workbook workbook,
            SheetValueStore store,
            string sheetName,
            int handle,
            int minColumn,
            int maxColumn,
            int minRow,
            int maxRow
        )
        {
            _workbook = workbook;
            _store = store;
            _sheetName = sheetName;
            _handle = handle;
            _minColumn = minColumn;
            _maxColumn = maxColumn;
            _minRow = minRow;
            _maxRow = maxRow;
            _column = minColumn;
            _row = minRow;
            _current = default;
        }

        public readonly ComputedValue Current => _current;

        readonly object System.Collections.IEnumerator.Current => _current;

        public bool MoveNext()
        {
            if (_column > _maxColumn)
            {
                return false;
            }

            // Hit path inline (matches the raw dense-read loop): a memoized cell is a lock-free store read with
            // no re-entry into GetCellValueDense. Only a MISS routes back through the workbook for the cycle
            // guard + on-demand evaluation + memoization — identical semantics, just no per-cell framing on hits.
            if (!_store.TryGetDense(_handle, _column, _row, out _current))
            {
                _current = _workbook.GetCellValueDense(_handle, _sheetName, _column, _row);
            }

            if (++_row > _maxRow)
            {
                _row = _minRow;
                _column++;
            }

            return true;
        }

        public void Reset()
        {
            _column = _minColumn;
            _row = _minRow;
            _current = default;
        }

        public readonly void Dispose() { }
    }
}
