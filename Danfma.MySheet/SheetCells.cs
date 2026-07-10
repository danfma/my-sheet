using Danfma.MySheet.Expressions;

namespace Danfma.MySheet;

/// <summary>
/// The numeric addresses of a sheet's populated canonical-A1 cells, in insertion order — the
/// allocation-free half of a bulk-extraction pipeline (pair each address with
/// <see cref="SheetValueReader.GetValue"/>, and render ids on demand with
/// <see cref="CellRef.TryFormat"/>). Cells stored under non-A1 ids (rare host-chosen keys) have no
/// numeric address and are not listed here — enumerate <see cref="Sheet.Keys"/> or
/// <see cref="Sheet.EnumerateCells"/> to see them. <c>foreach</c> binds to the struct enumerator
/// directly: no interface boxing, no per-item allocation.
/// </summary>
public readonly struct CellAddressCollection
{
    private readonly Dictionary<(int Column, int Row), Expression> _dense;

    internal CellAddressCollection(Dictionary<(int Column, int Row), Expression> dense) =>
        _dense = dense;

    /// <summary>Number of populated canonical-A1 cells.</summary>
    public int Count => _dense.Count;

    public Enumerator GetEnumerator() => new(_dense.GetEnumerator());

    public struct Enumerator
    {
        private Dictionary<(int Column, int Row), Expression>.Enumerator _inner;

        internal Enumerator(Dictionary<(int Column, int Row), Expression>.Enumerator inner) =>
            _inner = inner;

        public readonly (int Column, int Row) Current => _inner.Current.Key;

        public bool MoveNext() => _inner.MoveNext();
    }
}

/// <summary>
/// A sheet's populated cells as <c>(Id, Column, Row)</c>, in insertion order — the convenience form
/// for consumers that need the id text anyway (e.g. as a JSON field): the canonical id is derived
/// once per cell by the library instead of being rebuilt by the caller, and the numeric address
/// comes for free for <see cref="SheetValueReader.GetValue"/>. Note the id string IS allocated per
/// cell; for a fully allocation-free pipeline use <see cref="Sheet.CellAddresses"/> +
/// <see cref="CellRef.TryFormat"/>. Cells stored under non-A1 ids are included with
/// <c>Column = 0, Row = 0</c> (they have no numeric address — read them via
/// <see cref="Workbook.GetCellValue(string,string)"/>). <c>foreach</c> binds to the struct
/// enumerator directly: no interface boxing.
/// </summary>
public readonly struct SheetCellRefCollection
{
    private readonly Dictionary<(int Column, int Row), Expression> _dense;
    private readonly Dictionary<string, Expression>? _overflow;

    internal SheetCellRefCollection(
        Dictionary<(int Column, int Row), Expression> dense,
        Dictionary<string, Expression>? overflow
    )
    {
        _dense = dense;
        _overflow = overflow;
    }

    /// <summary>Number of populated cells (canonical and overflow).</summary>
    public int Count => _dense.Count + (_overflow?.Count ?? 0);

    public Enumerator GetEnumerator() => new(_dense, _overflow);

    public struct Enumerator
    {
        private Dictionary<(int Column, int Row), Expression>.Enumerator _dense;
        private Dictionary<string, Expression>.Enumerator _overflow;
        private readonly bool _hasOverflow;
        private bool _inOverflow;

        internal Enumerator(
            Dictionary<(int Column, int Row), Expression> dense,
            Dictionary<string, Expression>? overflow
        )
        {
            _dense = dense.GetEnumerator();
            _hasOverflow = overflow is not null;
            _overflow = overflow?.GetEnumerator() ?? default;
            _inOverflow = false;
            Current = default!;
        }

        public (string Id, int Column, int Row) Current { get; private set; }

        public bool MoveNext()
        {
            if (!_inOverflow)
            {
                if (_dense.MoveNext())
                {
                    var (column, row) = _dense.Current.Key;

                    Current = (new CellAddress(column, row).ToId(), column, row);

                    return true;
                }

                _inOverflow = true;
            }

            if (_hasOverflow && _overflow.MoveNext())
            {
                Current = (_overflow.Current.Key, 0, 0);

                return true;
            }

            return false;
        }
    }
}
