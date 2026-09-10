using System.Runtime.CompilerServices;
using Danfma.MySheet.Parsing;
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

    // === Validation (invoked by the registry, never by the constructor) ==================================

    private const int MaxNameLength = 255;

    /// <summary>
    /// Validates a table name against Excel's documented rule ("Rename an Excel table"): it must start with a
    /// letter or underscore, contain only letters, digits, '.' or '_', be at most 255 characters, and must
    /// not be one of the reserved single letters <c>C</c>/<c>R</c> (either case), a cell reference in A1 form
    /// (inside Excel's grid — <c>Tabela1</c> and <c>Table1</c>, Excel's own default names, are NOT cells
    /// because their letter runs exceed three characters) or in R1C1 form, or a boolean literal. The last is
    /// a MySheet addition: the tokenizer reads <c>TRUE</c>/<c>FALSE</c> as booleans, so such a table could
    /// never be reached from a formula. A leading backslash, which Excel permits, is rejected for the same
    /// reason — the tokenizer never reads one into an identifier. Throws <see cref="ArgumentException"/>.
    /// </summary>
    internal static void ValidateName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A table name cannot be empty.", nameof(name));
        }

        if (name.Length > MaxNameLength)
        {
            throw new ArgumentException(
                $"'{name[..16]}…' is not a valid table name: it is {name.Length} characters long and the "
                    + $"limit is {MaxNameLength}.",
                nameof(name)
            );
        }

        if (!IsValidName(name))
        {
            throw new ArgumentException(
                $"'{name}' is not a valid table name: it must start with a letter or underscore, contain "
                    + "only letters, digits, '.' or '_', and must not be \"C\" or \"R\", look like a cell "
                    + "reference (e.g. \"A1\" or \"R1C1\") or be a boolean literal.",
                nameof(name)
            );
        }
    }

    private static bool IsValidName(string name)
    {
        if (!char.IsLetter(name[0]) && name[0] != '_')
        {
            return false;
        }

        foreach (var c in name)
        {
            if (!char.IsLetterOrDigit(c) && c is not ('_' or '.'))
            {
                return false;
            }
        }

        if (name.Length == 1 && name[0] is 'C' or 'c' or 'R' or 'r')
        {
            return false;
        }

        return !Parser.IsExcelGridCellReference(name)
            && !IsR1C1Shape(name)
            && !string.Equals(name, "TRUE", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(name, "FALSE", StringComparison.OrdinalIgnoreCase);
    }

    // R<digits>C<digits>, either case: "R1C1", "r2c3".
    private static bool IsR1C1Shape(string name)
    {
        if (name.Length < 4 || name[0] is not ('R' or 'r'))
        {
            return false;
        }

        var i = 1;
        var rowDigits = 0;
        while (i < name.Length && char.IsAsciiDigit(name[i]))
        {
            i++;
            rowDigits++;
        }

        if (rowDigits == 0 || i == name.Length || name[i] is not ('C' or 'c'))
        {
            return false;
        }

        i++;
        var columnDigits = 0;
        while (i < name.Length && char.IsAsciiDigit(name[i]))
        {
            i++;
            columnDigits++;
        }

        return columnDigits > 0 && i == name.Length;
    }

    /// <summary>
    /// Validates the record's members at the registration boundary — the name rule, a non-blank sheet name,
    /// 1-based geometry, a range tall enough for the header and totals rows it claims (a range EXACTLY that
    /// tall, i.e. zero data rows, is legal), at least one column, and column names that are non-blank and
    /// unique ignoring case (Excel resolves <c>Table1[col]</c> to <c>Table1[Col]</c>). Column names are
    /// human-authored and are NOT subject to the name rule. Throws <see cref="ArgumentOutOfRangeException"/>
    /// for a row or column below 1 and <see cref="ArgumentException"/> for everything else.
    /// </summary>
    internal void Validate()
    {
        ValidateName(Name);

        ArgumentNullException.ThrowIfNull(SheetName);

        if (string.IsNullOrWhiteSpace(SheetName))
        {
            throw new ArgumentException(
                $"Table '{Name}': the sheet name cannot be empty.",
                nameof(SheetName)
            );
        }

        if (FirstRow < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(FirstRow), FirstRow, "Rows are 1-based.");
        }

        if (FirstColumn < 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(FirstColumn),
                FirstColumn,
                "Columns are 1-based."
            );
        }

        if (LastRow < FirstRow)
        {
            throw new ArgumentException(
                $"Table '{Name}': {nameof(LastRow)} ({LastRow}) is before {nameof(FirstRow)} ({FirstRow}).",
                nameof(LastRow)
            );
        }

        var reservedRows = (HasHeaderRow ? 1 : 0) + (HasTotalsRow ? 1 : 0);
        var rowSpan = LastRow - FirstRow + 1;
        if (rowSpan < reservedRows)
        {
            throw new ArgumentException(
                $"Table '{Name}': the range spans {rowSpan} row(s) but {nameof(HasHeaderRow)}={HasHeaderRow} "
                    + $"and {nameof(HasTotalsRow)}={HasTotalsRow} need {reservedRows}.",
                nameof(LastRow)
            );
        }

        ArgumentNullException.ThrowIfNull(ColumnNames);

        if (ColumnNames.Count == 0)
        {
            throw new ArgumentException(
                $"Table '{Name}': a table needs at least one column.",
                nameof(ColumnNames)
            );
        }

        var seen = new Dictionary<string, int>(ColumnNames.Count, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < ColumnNames.Count; i++)
        {
            var columnName = ColumnNames[i];

            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentException(
                    $"Table '{Name}': the column name at index {i} is empty.",
                    nameof(ColumnNames)
                );
            }

            if (!seen.TryAdd(columnName, i))
            {
                throw new ArgumentException(
                    $"Table '{Name}': the column name '{columnName}' at index {i} repeats the one at index "
                        + $"{seen[columnName]} (column names are compared ignoring case).",
                    nameof(ColumnNames)
                );
            }
        }
    }
}
