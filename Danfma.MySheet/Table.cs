using System.Runtime.CompilerServices;
using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// An Excel table (a "ListObject"): a named, sheet-anchored rectangle with named columns. The geometry
/// mirrors the xlsx <c>&lt;table&gt;</c> element exactly — <see cref="FirstRow"/>/<see cref="LastRow"/> span
/// the WHOLE <c>ref</c> (header row through totals row, both included when present) and
/// <see cref="FirstColumn"/> is the 1-based leftmost sheet column — so the header, data and totals rows are
/// DERIVED and can never contradict each other, and the last column follows from
/// <see cref="ColumnNames"/>. Value semantics only: validation lives in <see cref="Validate"/> and runs at the
/// registration boundary, never in the constructor, because MemoryPack materializes the record through the
/// same constructor and a hand-corrupted file must fail at the API, not during deserialization.
/// </summary>
/// <param name="Name">The table name, validated by <see cref="ValidateName"/>.</param>
/// <param name="SheetName">The sheet the table lives on.</param>
/// <param name="FirstRow">The 1-based first row of the whole table range (the header row when there is one).</param>
/// <param name="LastRow">The 1-based last row of the whole table range (the totals row when there is one).</param>
/// <param name="FirstColumn">The 1-based leftmost sheet column.</param>
/// <param name="HasHeaderRow">Whether <see cref="FirstRow"/> is a header row.</param>
/// <param name="HasTotalsRow">Whether <see cref="LastRow"/> is a totals row.</param>
/// <param name="ColumnNames">The column names in sheet order; one per column of the range.</param>
[MemoryPackable]
public sealed partial record Table(
    string Name,
    string SheetName,
    int FirstRow,
    int LastRow,
    int FirstColumn,
    bool HasHeaderRow,
    bool HasTotalsRow,
    IReadOnlyList<string> ColumnNames
)
{
    // === Derived geometry (runtime-only — never on the wire) =============================================

    /// <summary>The number of columns, i.e. <c>ColumnNames.Count</c>.</summary>
    [MemoryPackIgnore]
    public int ColumnCount => ColumnNames.Count;

    /// <summary>The 1-based rightmost sheet column.</summary>
    [MemoryPackIgnore]
    public int LastColumn => FirstColumn + ColumnNames.Count - 1;

    /// <summary>The 1-based header row, or <c>null</c> when the table has none.</summary>
    [MemoryPackIgnore]
    public int? HeaderRow => HasHeaderRow ? FirstRow : null;

    /// <summary>The 1-based totals row, or <c>null</c> when the table has none.</summary>
    [MemoryPackIgnore]
    public int? TotalsRow => HasTotalsRow ? LastRow : null;

    /// <summary>The 1-based first data row (the row after the header, or <see cref="FirstRow"/>).</summary>
    [MemoryPackIgnore]
    public int FirstDataRow => FirstRow + (HasHeaderRow ? 1 : 0);

    /// <summary>The 1-based last data row (the row before the totals, or <see cref="LastRow"/>).</summary>
    [MemoryPackIgnore]
    public int LastDataRow => LastRow - (HasTotalsRow ? 1 : 0);

    /// <summary>
    /// The number of data rows, clamped at 0: a header-only table (<c>ref="A1:A1"</c> with a header row) is a
    /// legal model state with no data, not a negative count.
    /// </summary>
    [MemoryPackIgnore]
    public int DataRowCount => Math.Max(0, LastDataRow - FirstDataRow + 1);

    // === Column lookup ===================================================================================

    // The lazily built name -> ordinal map is memoized OFF the record, keyed by reference identity. A record's
    // synthesized copy constructor copies every instance field — a memo field would travel with
    // `table with { ColumnNames = ... }` and answer for the OLD names (measured: a warmed table copied with a
    // third column reported that column absent), and it would take part in the synthesized Equals/GetHashCode,
    // so a table's hash would change on its first lookup. With the memo here, a `with` copy is a new key that
    // starts cold, equality stays over the eight members, and the entry dies with the table. Built once per
    // table instance; GetValue is thread-safe (a losing racer's map is dropped, and the maps are identical).
    private static readonly ConditionalWeakTable<Table, Dictionary<string, int>> ColumnIndexes =
        new();

    /// <summary>
    /// Resolves a column name (case-insensitively, as Excel resolves <c>Table1[col]</c> to <c>Table1[Col]</c>)
    /// to its 0-based ordinal in <see cref="ColumnNames"/>. Returns <c>false</c> (index 0) for an unknown name.
    /// </summary>
    public bool TryGetColumnIndex(string columnName, out int index) =>
        ColumnIndexes
            .GetValue(this, static table => table.BuildColumnIndex())
            .TryGetValue(columnName, out index);

    private Dictionary<string, int> BuildColumnIndex()
    {
        var map = new Dictionary<string, int>(ColumnNames.Count, StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < ColumnNames.Count; i++)
        {
            map.TryAdd(ColumnNames[i], i);
        }

        return map;
    }

    // === Sheet coordinates ===============================================================================

    /// <summary>The 1-based sheet column of the 0-based table column <paramref name="columnIndex"/>.</summary>
    public int SheetColumnAt(int columnIndex) => FirstColumn + columnIndex;

    /// <summary>
    /// The [#Data] rows of the named column as sheet coordinates — the whole surface the reference-semantics
    /// phase needs to turn <c>Tabela1[Valor]</c> into a concrete range without re-deriving geometry. Returns
    /// <c>false</c> (all outputs 0) when the name is unknown OR the table has no data rows: the caller owns the
    /// <c>#REF!</c> decision and this is the one place that knows the table is empty.
    /// </summary>
    public bool TryGetColumnRange(
        string columnName,
        out int sheetColumn,
        out int firstRow,
        out int lastRow
    )
    {
        if (DataRowCount == 0 || !TryGetColumnIndex(columnName, out var index))
        {
            sheetColumn = 0;
            firstRow = 0;
            lastRow = 0;
            return false;
        }

        sheetColumn = SheetColumnAt(index);
        firstRow = FirstDataRow;
        lastRow = LastDataRow;
        return true;
    }
}
