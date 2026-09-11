using System.Runtime.CompilerServices;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet;

/// <summary>
/// What <see cref="Table.GetRegion"/> found. Three outcomes, not two, and <see cref="Empty"/> is
/// deliberately NOT folded into <see cref="Absent"/>: both answer <c>#REF!</c> today, but they answer it for
/// different reasons and only one of them is a recorded DIVERGENCE from the oracle. Measured on
/// Aspose.Cells 26.6.0 (2026-09-11), a header-only table's data band is an EMPTY reference there —
/// <c>SUM</c> 0, <c>COUNT</c> 0, <c>COUNTA</c> 0, <c>ROWS</c> 0, <c>COLUMNS</c> 1, <c>ISREF</c> TRUE,
/// <c>SUBTOTAL(9)</c> 0, <c>AVERAGE</c> <c>#DIV/0!</c>, <c>INDEX(…,1,1)</c> <c>#REF!</c> — and this engine
/// has no zero-extent reference node to answer with, so the resolver maps <see cref="Empty"/> to
/// <c>#REF!</c> in ONE arm. Keeping the outcome separate is what makes reopening that ruling a single edit
/// per consumer instead of a re-plumbing of this primitive.
/// </summary>
internal enum TableRegionOutcome : byte
{
    /// <summary>A real rectangle: the four out parameters are 1-based sheet coordinates.</summary>
    Resolved = 0,

    /// <summary>
    /// The area exists but spans zero rows — a header-only table's <c>[#Data]</c>, or either pair left with
    /// neither the row it names nor a data row. The out parameters are 0, and handing back an inverted
    /// rectangle instead is not an option, because <c>RangeReference.GetBounds</c> normalizes min/max
    /// (measured: a range built from <c>B2</c>..<c>B1</c> reports <c>TopRow</c> 1 and <c>ROWS</c> 2, and
    /// <c>SUM</c> over it reads both cells), so a zero-data-row table handed back as
    /// <c>(top 2, bottom 1)</c> would silently read the HEADER row.
    /// </summary>
    Empty,

    /// <summary>
    /// The area or the column does not exist on this table: <c>[#Headers]</c> without a header row,
    /// <c>[#Totals]</c> without a totals row, an unknown column name, an area value off the wire that no
    /// member names, or a <c>ref</c> that cannot be a rectangle at all (no columns, or <c>LastRow</c> before
    /// <c>FirstRow</c> — both rejected by <see cref="Table.Validate"/> but reachable through
    /// deserialization). The out parameters are 0.
    /// </summary>
    Absent,
}

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
    /// The [#Data] rows of the named column as sheet coordinates: the <see cref="TableArea.Data"/> case of
    /// <see cref="GetRegion"/>, which owns the geometry for all six areas. Returns <c>false</c> (all outputs
    /// 0) when the name is unknown OR the table has no data rows, and the caller owns the <c>#REF!</c>
    /// decision. Telling those two reasons apart needs <see cref="GetRegion"/> and its
    /// <see cref="TableRegionOutcome"/>, which are internal to this assembly — outside it, <c>false</c> is
    /// one outcome.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="columnName"/> is <c>null</c>. It always throws now; before the region primitive
    /// existed, a table with zero data rows short-circuited on the row count and answered <c>false</c>
    /// without looking at the name.
    /// </exception>
    public bool TryGetColumnRange(
        string columnName,
        out int sheetColumn,
        out int firstRow,
        out int lastRow
    )
    {
        // GetRegion reads a null column as "every column of the band" on purpose (that is T[#Data]), so the
        // non-nullable contract of this overload has to be enforced here rather than inherited from it.
        ArgumentNullException.ThrowIfNull(columnName);

        var outcome = GetRegion(
            columnName,
            TableArea.Data,
            out sheetColumn,
            out firstRow,
            out _,
            out lastRow
        );
        return outcome == TableRegionOutcome.Resolved;
    }

    /// <summary>
    /// The ONE geometry primitive: the sheet rectangle a structured reference names, as 1-based
    /// <paramref name="left"/>/<paramref name="top"/>/<paramref name="right"/>/<paramref name="bottom"/>.
    /// The row band is computed first and the column narrowing applied to it, which is how every specifier
    /// form falls out of one function: <c>T[Col]</c> is <see cref="TableArea.Data"/> plus a column,
    /// <c>T[[#Data],[Col]]</c> is the identical node, and
    /// <c>T[[#Headers],[Col]]</c>/<c>T[[#Totals],[Col]]</c>/<c>T[[#All],[Col]]</c> come for free. (The column
    /// NAME is checked before either, so a typo reports <see cref="TableRegionOutcome.Absent"/> rather than
    /// hiding behind a band that happens to be empty.)
    /// <para>
    /// The bands, all measured on Aspose.Cells 26.6.0 (2026-09-11) over <c>Data!Tabela1</c> = <c>A1:C4</c>
    /// with and without a totals row, PLAIN and array-entered (the two modes agreed on every row):
    /// <see cref="TableArea.All"/> is the whole <c>ref</c> (<c>A1:C4</c> / <c>A1:C5</c>),
    /// <see cref="TableArea.Data"/> the data rows (<c>A2:C4</c> either way),
    /// <see cref="TableArea.Headers"/> the header row (<c>A1:C1</c>), <see cref="TableArea.Totals"/> the
    /// totals row (<c>A5:C5</c>), <see cref="TableArea.HeadersAndData"/> <c>TopRow..LastDataRow</c>
    /// (<c>A1:C4</c> either way) and <see cref="TableArea.DataAndTotals"/> <c>FirstDataRow..BottomRow</c>
    /// (<c>A2:C4</c> / <c>A2:C5</c>).
    /// </para>
    /// <para>
    /// Of the two areas that can be absent, only the SINGLETONS are absent because the row they name is
    /// missing. A PAIR in that position SHRINKS to the rows it still has — measured,
    /// <c>SUM(T[[#Data],[#Totals]])</c> over the A1:C4 fixture with no totals row is the data body, 66, not
    /// <c>#REF!</c>, and over a header-LESS table (a different fixture: two data rows at A2:C3 summing 55)
    /// <c>SUM(T[[#Headers],[#Data]])</c> is likewise the data, 55, while <c>SUM(T[#Headers])</c> there is
    /// <c>#REF!</c>. Copying a singleton's absent arm into a pair by analogy is the mistake that answers
    /// <c>#REF!</c> where Excel answers the data, in the very shape users write to mean "the table without
    /// its header". (A pair CAN still end up <see cref="TableRegionOutcome.Empty"/>, and so <c>#REF!</c> at
    /// the resolver, when it is left with no rows at all — a header-only table's
    /// <c>[[#Data],[#Totals]]</c>.)
    /// </para>
    /// </summary>
    internal TableRegionOutcome GetRegion(
        string? columnName,
        TableArea area,
        out int left,
        out int top,
        out int right,
        out int bottom
    )
    {
        left = 0;
        top = 0;
        right = 0;
        bottom = 0;

        // A ref that cannot be a rectangle at all has no areas: no columns (ColumnNames empty, so
        // LastColumn == FirstColumn - 1) or LastRow before FirstRow. Validate rejects both at the
        // registration boundary, but MemoryPack materializes the record through the same constructor and
        // never runs Validate, so a hand-corrupted file can carry one — and it MUST be caught here.
        // Measured: over a zero-column table this method used to answer Resolved with (left 1, right 0),
        // TryResolveRange then built CellAddress(0, 4).ToId() == "4", and SUM/ROWS/COLUMNS/Evaluate all THREW
        // FormatException "Invalid cell reference '4'" out of RangeReference.GetBounds — the one thing
        // Evaluate's documented contract says it never does. Absent rather than Empty: an impossible ref is
        // the same class of input as an area byte no member names, and Empty carries a live ruling that
        // corrupt geometry must not be dragged into.
        if (LastRow < FirstRow || LastColumn < FirstColumn)
        {
            return TableRegionOutcome.Absent;
        }

        // The column is resolved BEFORE the band: a name that does not exist is the formula's own bug and
        // must report Absent whatever the band turns out to be, never hide behind the data-dependent Empty.
        var index = -1;
        if (columnName is not null && !TryGetColumnIndex(columnName, out index))
        {
            return TableRegionOutcome.Absent;
        }

        int bandTop;
        int bandBottom;
        switch (area)
        {
            case TableArea.All:
                (bandTop, bandBottom) = (FirstRow, LastRow);
                break;
            case TableArea.Data:
                (bandTop, bandBottom) = (FirstDataRow, LastDataRow);
                break;
            case TableArea.Headers when HasHeaderRow:
                (bandTop, bandBottom) = (FirstRow, FirstRow);
                break;
            case TableArea.Totals when HasTotalsRow:
                (bandTop, bandBottom) = (LastRow, LastRow);
                break;
            case TableArea.HeadersAndData:
                (bandTop, bandBottom) = (FirstRow, LastDataRow);
                break;
            case TableArea.DataAndTotals:
                (bandTop, bandBottom) = (FirstDataRow, LastRow);
                break;
            // [#Headers] with no header row, [#Totals] with no totals row, and an area value off the wire
            // that no member names (Validate does not run on deserialization).
            default:
                return TableRegionOutcome.Absent;
        }

        // One check for all six bands. Together with the impossible-ref guard above it, this is what makes
        // "no inverted rectangle leaves this method" true on BOTH axes — see TableRegionOutcome.Empty for the
        // measurement that makes it mandatory rather than stylistic.
        if (bandBottom < bandTop)
        {
            return TableRegionOutcome.Empty;
        }

        (left, right) =
            index < 0 ? (FirstColumn, LastColumn) : (SheetColumnAt(index), SheetColumnAt(index));
        (top, bottom) = (bandTop, bandBottom);
        return TableRegionOutcome.Resolved;
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
