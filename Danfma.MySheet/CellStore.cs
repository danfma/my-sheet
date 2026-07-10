using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Danfma.MySheet.Expressions;
using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// The in-memory backing store for a <see cref="Sheet"/>'s populated cells. Canonical A1 ids
/// (<c>A1</c>, <c>AA100</c> — uppercase column letters, no leading zero, in range) are held in a numeric
/// <c>(column,row)</c> dictionary — the string-hygiene 3.x lever: at K1 scale the old
/// <c>Dictionary&lt;string,Expression&gt;</c> retained ~52.9MB of duplicate id strings, the numeric key
/// collapses that to ~18MB. Any id that is NOT canonical A1 (a host-supplied non-address key, a lowercase
/// or padded form) rounds through an overflow <c>Dictionary&lt;string,Expression&gt;</c> so the exact key
/// string is preserved on the public surface — the same split the value store already uses.
///
/// <para>The public <see cref="Sheet"/> surface (the string indexer, <see cref="Sheet.Keys"/>,
/// enumeration, <see cref="Sheet.TryGetValue"/>, <see cref="Sheet.ContainsKey"/>) reads through this type,
/// deriving the A1 id on demand for enumeration (a cold path). The wire format is unchanged: the
/// <see cref="CellStoreFormatter"/> serializes and deserializes the historical
/// <c>Dictionary&lt;string,Expression&gt;</c> byte-for-byte (see that type for the ordering argument).</para>
/// </summary>
internal sealed class CellStore : IReadOnlyDictionary<string, Expression>
{
    // Canonical A1 cells, addressed numerically. A .NET Dictionary preserves insertion order until a
    // removal reuses a slot; since the string store this replaces is filled by the same SetCell sequence,
    // the two enumerate in the same order — the byte-identity argument for the wire (CellStoreFormatter).
    private readonly Dictionary<(int Column, int Row), Expression> _dense;

    // Non-A1 ids (a host storing a cell under an arbitrary key). Lazily created — dormant in normal use.
    private Dictionary<string, Expression>? _overflow;

    public CellStore() => _dense = new();

    // The formatter presizes the dense map for the incoming A1 cells (overflow stays lazy).
    internal CellStore(int denseCapacity) => _dense = new(denseCapacity);

    // Bulk loaders (the .xlsx reader) reserve capacity up front to avoid rehash cascades while cells
    // stream in. Dictionary.EnsureCapacity never reorders existing entries, so the wire byte-identity
    // argument above is unaffected.
    internal void EnsureDenseCapacity(int capacity) => _dense.EnsureCapacity(capacity);

    /// <summary>
    /// Reads the numeric address out of a canonical A1 id WITHOUT allocating, returning <c>true</c> only
    /// when the id is EXACTLY what <see cref="CellAddress.ToId"/> would produce for <c>(column,row)</c>:
    /// uppercase column letters (no <c>$</c>, no lowercase), a row with no leading zero, both within
    /// <see cref="int"/> range. That round-trip guarantee is what lets enumeration re-derive the exact key.
    /// A lenient parse (e.g. lowercase, padding) would change the key on the way back, so such ids go to the
    /// overflow store instead.
    /// </summary>
    internal static bool TryParseCanonical(string id, out int column, out int row)
    {
        column = 0;
        row = 0;

        var length = id.Length;
        if (length == 0)
        {
            return false;
        }

        var i = 0;
        long col = 0;
        while (i < length)
        {
            var c = id[i];
            if (c is < 'A' or > 'Z')
            {
                break;
            }

            col = col * 26 + (c - 'A' + 1);
            if (col > int.MaxValue)
            {
                return false;
            }

            i++;
        }

        // Need at least one letter and at least one following digit.
        if (i == 0 || i == length)
        {
            return false;
        }

        // A leading zero (or a bare "0") is not what ToId emits — keep the exact string in overflow.
        if (id[i] == '0')
        {
            return false;
        }

        long parsedRow = 0;
        for (; i < length; i++)
        {
            var c = id[i];
            if (c is < '0' or > '9')
            {
                return false;
            }

            parsedRow = parsedRow * 10 + (c - '0');
            if (parsedRow > int.MaxValue)
            {
                return false;
            }
        }

        column = (int)col;
        row = (int)parsedRow;
        return true;
    }

    // The single write path: canonical A1 goes to the dense map, everything else to the overflow map.
    // Reports whether a NEW dense (A1) cell was inserted — the only case the structural index must learn —
    // along with its coordinates (valid only when the flag is set).
    internal void Set(
        string id,
        Expression expr,
        out bool addedDenseCell,
        out int column,
        out int row
    )
    {
        if (TryParseCanonical(id, out column, out row))
        {
            ref var slot = ref CollectionsMarshal.GetValueRefOrAddDefault(
                _dense,
                (column, row),
                out var existed
            );
            slot = expr;
            addedDenseCell = !existed;
            return;
        }

        (_overflow ??= new())[id] = expr;
        addedDenseCell = false;
    }

    // The delete path, symmetric with Set: reports whether the removed cell was a dense (A1) one (so the
    // structural index can drop it), with its coordinates.
    internal bool Remove(string id, out bool wasDenseCell, out int column, out int row)
    {
        if (TryParseCanonical(id, out column, out row))
        {
            wasDenseCell = _dense.Remove((column, row));
            return wasDenseCell;
        }

        wasDenseCell = false;
        return _overflow is not null && _overflow.Remove(id);
    }

    // The dense (A1) coordinates, for the structural-index build — it consumes numeric addresses directly,
    // so the build no longer re-derives an id string per cell only to re-parse it (the overflow ids are not
    // A1 and were skipped by the old parse anyway).
    internal CellAddressCollection DenseAddresses => new(_dense);

    internal SheetCellRefCollection AddressedEntries => new(_dense, _overflow);

    public int Count => _dense.Count + (_overflow?.Count ?? 0);

    public bool ContainsKey(string key)
    {
        if (TryParseCanonical(key, out var column, out var row))
        {
            return _dense.ContainsKey((column, row));
        }

        return _overflow is not null && _overflow.ContainsKey(key);
    }

    public bool TryGetValue(string key, [MaybeNullWhen(false)] out Expression value)
    {
        if (TryParseCanonical(key, out var column, out var row))
        {
            return _dense.TryGetValue((column, row), out value);
        }

        if (_overflow is not null)
        {
            return _overflow.TryGetValue(key, out value);
        }

        value = null;
        return false;
    }

    public Expression this[string key] =>
        TryGetValue(key, out var value)
            ? value
            : throw new KeyNotFoundException($"The cell '{key}' is not populated.");

    /// <summary>
    /// Numeric-address read of the dense (A1) map — for a caller that already has <c>(column, row)</c> in
    /// hand from a numeric scan (e.g. a range walk or the structural index) and would otherwise pay a
    /// <see cref="CellAddress.ToId"/> build plus a re-parse just to ask "is this cell populated / what
    /// expression does it hold". Non-A1 overflow cells are not addressable this way (as with every other
    /// numeric accessor on this store) — a caller that needs those still goes through the string indexer.
    /// </summary>
    internal bool TryGetDense(int column, int row, [MaybeNullWhen(false)] out Expression value) =>
        _dense.TryGetValue((column, row), out value);

    // The A1 ids are re-derived on demand (a cold path — enumeration/Keys, never per formula read).
    public IEnumerable<string> Keys
    {
        get
        {
            foreach (var (column, row) in _dense.Keys)
            {
                yield return new CellAddress(column, row).ToId();
            }

            if (_overflow is not null)
            {
                foreach (var id in _overflow.Keys)
                {
                    yield return id;
                }
            }
        }
    }

    public IEnumerable<Expression> Values
    {
        get
        {
            foreach (var value in _dense.Values)
            {
                yield return value;
            }

            if (_overflow is not null)
            {
                foreach (var value in _overflow.Values)
                {
                    yield return value;
                }
            }
        }
    }

    public IEnumerator<KeyValuePair<string, Expression>> GetEnumerator()
    {
        foreach (var (address, value) in _dense)
        {
            yield return new KeyValuePair<string, Expression>(
                new CellAddress(address.Column, address.Row).ToId(),
                value
            );
        }

        if (_overflow is not null)
        {
            foreach (var pair in _overflow)
            {
                yield return pair;
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

/// <summary>
/// The per-member MemoryPack formatter that keeps <see cref="CellStore"/>'s wire format byte-identical to the
/// historical <c>Dictionary&lt;string,Expression&gt;</c> the cell store used to be. Serialize materializes the
/// string-keyed view in enumeration order and hands it to the SAME built-in dictionary formatter the generated
/// <c>Sheet</c> code used before (so the bytes are identical); Deserialize reads that dictionary and splits it
/// back into the numeric dense map + non-A1 overflow. Only Save/Load pay the string round trip (once), never a
/// per-cell read.
///
/// <para><b>Ordering.</b> For a workbook of only A1 cells (every real file) the dense map enumerates in the same
/// order the old string map did — both are filled by the identical SetCell sequence and a .NET Dictionary's
/// layout is fixed by its operation order, not its keys — so the reconstructed string view matches the original
/// byte-for-byte. When non-A1 overflow ids are present the A1 entries keep their relative order and the overflow
/// entries follow; that stays stable across new→new round trips, only the exact interleave with a pre-existing
/// mixed file is not reproduced (non-A1 keys are a host edge case, absent from every fixture/real workbook).</para>
/// </summary>
internal sealed class CellStoreFormatter : MemoryPackFormatter<CellStore>
{
    public static readonly CellStoreFormatter Default = new();

    static CellStoreFormatter()
    {
        // The store's member type is no longer Dictionary<string, Expression>, so the generated Sheet
        // formatter stops registering that dictionary formatter — yet Serialize/Deserialize delegate to it for
        // the byte-identical wire. Register it here so the delegation never depends on another member (e.g.
        // Workbook.DefinedNames) happening to carry the same type.
        if (!MemoryPackFormatterProvider.IsRegistered<Dictionary<string, Expression>>())
        {
            MemoryPackFormatterProvider.Register(
                new global::MemoryPack.Formatters.DictionaryFormatter<string, Expression>()
            );
        }
    }

    public override void Serialize<TBufferWriter>(
        ref MemoryPackWriter<TBufferWriter> writer,
        scoped ref CellStore? value
    )
    {
        if (value is null)
        {
            // Mirror a null Dictionary member (never happens in practice — the field is always initialized).
            Dictionary<string, Expression>? nothing = null;
            writer.WriteValue(nothing);
            return;
        }

        // Rebuild the historical string-keyed map in enumeration order, then delegate to the exact same
        // dictionary formatter the generated Sheet code called before — byte-identical output.
        var wire = new Dictionary<string, Expression>(value.Count);
        foreach (var pair in value)
        {
            wire[pair.Key] = pair.Value;
        }

        writer.WriteValue(wire);
    }

    public override void Deserialize(ref MemoryPackReader reader, scoped ref CellStore? value)
    {
        Dictionary<string, Expression>? wire = null;
        reader.ReadValue(ref wire);

        if (wire is null)
        {
            value = null;
            return;
        }

        var store = new CellStore(wire.Count);
        foreach (var pair in wire)
        {
            store.Set(pair.Key, pair.Value, out _, out _, out _);
        }

        value = store;
    }
}

/// <summary>Applies <see cref="CellStoreFormatter"/> to a <see cref="CellStore"/> member.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
internal sealed class CellStoreFormatterAttribute
    : MemoryPackCustomFormatterAttribute<CellStoreFormatter, CellStore>
{
    public override CellStoreFormatter GetFormatter() => CellStoreFormatter.Default;
}
