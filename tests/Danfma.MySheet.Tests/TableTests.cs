using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Tests;

/// <summary>
/// The <see cref="Table"/> record on its own — the geometry it derives from the xlsx-shaped
/// <c>FirstRow/LastRow/FirstColumn + HasHeaderRow/HasTotalsRow</c> members, the column lookup, and the
/// six-area region primitive (<see cref="Table.GetRegion"/>) the reference-semantics phase resolves a
/// structured reference through. Nothing here touches a <see cref="Workbook"/>: the registry's pins live in
/// <c>TableRegistryTests</c>.
/// </summary>
public class TableTests
{
    // Data!C2:E9 with a header AND a totals row: header 2, data 3..8, totals 9 — the shape a ClosedXML-written
    // file reports as ref="C2:E9" totalsRowCount="1".
    private static Table WithTotals() =>
        new(
            "Tabela1",
            "Data",
            FirstRow: 2,
            LastRow: 9,
            FirstColumn: 3,
            HasHeaderRow: true,
            HasTotalsRow: true,
            ["Produto", "Qtd", "Total"]
        );

    // === Derived geometry ================================================================================

    [Test]
    public async Task Geometry_WithHeaderAndTotals_DerivesEveryRow()
    {
        var table = WithTotals();

        await Assert.That(table.ColumnCount).IsEqualTo(3);
        await Assert.That(table.LastColumn).IsEqualTo(5);
        await Assert.That(table.HeaderRow).IsEqualTo(2);
        await Assert.That(table.FirstDataRow).IsEqualTo(3);
        await Assert.That(table.LastDataRow).IsEqualTo(8);
        await Assert.That(table.DataRowCount).IsEqualTo(6);
        await Assert.That(table.TotalsRow).IsEqualTo(9);
    }

    // Without a header row the data starts ON FirstRow; without a totals row it ends ON LastRow — the flags
    // are the only thing that moves the data band, so HeaderRow/TotalsRow are null, not 0.
    [Test]
    public async Task Geometry_WithoutHeaderOrTotals_DataSpansTheWholeRef()
    {
        var table = new Table(
            "Raw",
            "Data",
            5,
            7,
            2,
            HasHeaderRow: false,
            HasTotalsRow: false,
            ["A"]
        );

        await Assert.That(table.HeaderRow).IsNull();
        await Assert.That(table.TotalsRow).IsNull();
        await Assert.That(table.FirstDataRow).IsEqualTo(5);
        await Assert.That(table.LastDataRow).IsEqualTo(7);
        await Assert.That(table.DataRowCount).IsEqualTo(3);
        await Assert.That(table.LastColumn).IsEqualTo(2);
    }

    // ref="A1:A1" with a header row is a LEGAL table with zero data rows: LastDataRow (0) < FirstDataRow (2),
    // and the count clamps at 0 instead of going negative.
    [Test]
    public async Task Geometry_HeaderOnly_ClampsDataRowCountAtZero()
    {
        var table = new Table(
            "Solo",
            "Data",
            1,
            1,
            1,
            HasHeaderRow: true,
            HasTotalsRow: false,
            ["A"]
        );

        await Assert.That(table.FirstDataRow).IsEqualTo(2);
        await Assert.That(table.LastDataRow).IsEqualTo(1);
        await Assert.That(table.DataRowCount).IsEqualTo(0);
    }

    // Header + totals over two rows is the other zero-data shape: header 1, totals 2, nothing between.
    [Test]
    public async Task Geometry_HeaderAndTotalsOnly_HasZeroDataRows()
    {
        var table = new Table(
            "Solo",
            "Data",
            1,
            2,
            1,
            HasHeaderRow: true,
            HasTotalsRow: true,
            ["A"]
        );

        await Assert.That(table.FirstDataRow).IsEqualTo(2);
        await Assert.That(table.LastDataRow).IsEqualTo(1);
        await Assert.That(table.DataRowCount).IsEqualTo(0);
        await Assert.That(table.TotalsRow).IsEqualTo(2);
    }

    // === Column lookup ===================================================================================

    [Test]
    public async Task TryGetColumnIndex_ResolvesEveryColumn_IgnoringCase()
    {
        var table = WithTotals();

        await Assert.That(table.TryGetColumnIndex("produto", out var produto)).IsTrue();
        await Assert.That(produto).IsEqualTo(0);
        await Assert.That(table.TryGetColumnIndex("QTD", out var qtd)).IsTrue();
        await Assert.That(qtd).IsEqualTo(1);
        await Assert.That(table.TryGetColumnIndex("Total", out var total)).IsTrue();
        await Assert.That(total).IsEqualTo(2);
    }

    [Test]
    public async Task TryGetColumnIndex_UnknownName_IsFalseAndZero()
    {
        var table = WithTotals();

        await Assert.That(table.TryGetColumnIndex("Imposto", out var index)).IsFalse();
        await Assert.That(index).IsEqualTo(0);
    }

    // The index is built over whatever list backs ColumnNames — a List<string> is what a Load produces.
    [Test]
    public async Task TryGetColumnIndex_WorksOverAListBackedColumnNames()
    {
        var table = WithTotals() with
        {
            ColumnNames = new List<string> { "X", "Y" },
        };

        await Assert.That(table.TryGetColumnIndex("y", out var y)).IsTrue();
        await Assert.That(y).IsEqualTo(1);
    }

    // A record's synthesized copy constructor copies EVERY instance field, so a memo stored on the record
    // would travel with a `with` copy and answer for the OLD ColumnNames. This is the flow the registry itself
    // uses (DefineTable snapshots via `with { ColumnNames = ... }`) and the natural host flow for resizing a
    // table after evaluation has warmed the memo — so the copy must resolve by the NEW names and report the
    // old ones absent, whatever it inherited.
    [Test]
    public async Task With_AfterTheIndexWasBuilt_ResolvesByTheNewColumnNames()
    {
        var original = new Table("Tabela", "Data", 1, 3, 1, true, false, ["a", "b"]);
        await Assert.That(original.TryGetColumnIndex("a", out _)).IsTrue(); // warms the memo

        var renamed = original with { ColumnNames = ["b", "a", "c"] };

        await Assert.That(renamed.TryGetColumnIndex("c", out var c)).IsTrue();
        await Assert.That(c).IsEqualTo(2);
        await Assert.That(renamed.TryGetColumnIndex("a", out var a)).IsTrue();
        await Assert.That(a).IsEqualTo(1);
        await Assert.That(renamed.TryGetColumnIndex("b", out var b)).IsTrue();
        await Assert.That(b).IsEqualTo(0);

        var replaced = original with { ColumnNames = ["x"] };

        await Assert.That(replaced.TryGetColumnIndex("a", out _)).IsFalse();
        await Assert.That(replaced.TryGetColumnIndex("b", out _)).IsFalse();
        await Assert.That(replaced.TryGetColumnIndex("x", out var x)).IsTrue();
        await Assert.That(x).IsEqualTo(0);

        // The original is untouched by the copies.
        await Assert.That(original.TryGetColumnIndex("c", out _)).IsFalse();
        await Assert.That(original.TryGetColumnIndex("b", out var originalB)).IsTrue();
        await Assert.That(originalB).IsEqualTo(1);
    }

    // The memo must not take part in the record's value semantics either: two tables built from the same
    // members are equal and hash alike BEFORE and AFTER one of them has answered a lookup. (ColumnNames
    // itself still compares by reference — that is the record default and is not what this pins.)
    [Test]
    public async Task TryGetColumnIndex_DoesNotChangeEqualityOrHashCode()
    {
        string[] columns = ["a", "b"];
        var x = new Table("Tabela", "Data", 1, 3, 1, true, false, columns);
        var y = new Table("Tabela", "Data", 1, 3, 1, true, false, columns);
        var hashBefore = x.GetHashCode();

        await Assert.That(x == y).IsTrue();
        await Assert.That(x.TryGetColumnIndex("a", out _)).IsTrue();

        await Assert.That(x == y).IsTrue();
        await Assert.That(x.GetHashCode()).IsEqualTo(hashBefore);
        await Assert.That(x.GetHashCode()).IsEqualTo(y.GetHashCode());
    }

    // === Sheet coordinates ===============================================================================

    [Test]
    public async Task SheetColumnAt_MapsTheZeroBasedOrdinalOntoTheSheet()
    {
        var table = WithTotals();

        await Assert.That(table.SheetColumnAt(0)).IsEqualTo(3);
        await Assert.That(table.SheetColumnAt(2)).IsEqualTo(5);
    }

    // [#Data] for a column: the sheet column plus the data band only — never the header or totals row.
    [Test]
    public async Task TryGetColumnRange_ReturnsTheDataBandOfThatColumn()
    {
        var table = WithTotals();

        var ok = table.TryGetColumnRange("qtd", out var column, out var firstRow, out var lastRow);

        await Assert.That(ok).IsTrue();
        await Assert.That(column).IsEqualTo(4);
        await Assert.That(firstRow).IsEqualTo(3);
        await Assert.That(lastRow).IsEqualTo(8);
    }

    [Test]
    public async Task TryGetColumnRange_UnknownColumn_IsFalse()
    {
        var table = WithTotals();

        var ok = table.TryGetColumnRange(
            "Imposto",
            out var column,
            out var firstRow,
            out var lastRow
        );

        await Assert.That(ok).IsFalse();
        await Assert.That((column, firstRow, lastRow)).IsEqualTo((0, 0, 0));
    }

    // A known column on a table with no data rows is the #REF! decision handed to the consumer: false, not
    // an inverted (firstRow > lastRow) range.
    [Test]
    public async Task TryGetColumnRange_KnownColumnButNoDataRows_IsFalse()
    {
        var table = new Table(
            "Solo",
            "Data",
            1,
            1,
            1,
            HasHeaderRow: true,
            HasTotalsRow: false,
            ["A"]
        );

        var ok = table.TryGetColumnRange("A", out _, out var firstRow, out var lastRow);

        await Assert.That(ok).IsFalse();
        await Assert.That((firstRow, lastRow)).IsEqualTo((0, 0));
    }

    // GetRegion reads a null column name as "every column of the band" — that is what T[#Data] means — so
    // this overload, whose parameter is not nullable, keeps its own guard instead of inheriting that. The
    // zero-data-row row is the one that MOVED: the old body was
    // `DataRowCount == 0 || !TryGetColumnIndex(columnName, …)`, which short-circuited on the row count and
    // answered false without ever looking at the name, so whether a null name threw depended on whether the
    // table had rows. A table WITH data threw either way, which is why it cannot be the only case here.
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task TryGetColumnRange_NullColumnName_Throws(bool hasDataRows)
    {
        var table = hasDataRows
            ? WithTotals()
            : new Table("Solo", "Data", 1, 1, 1, HasHeaderRow: true, HasTotalsRow: false, ["A"]);

        await Assert.That(table.DataRowCount).IsEqualTo(hasDataRows ? 6 : 0);
        await Assert
            .That(() => table.TryGetColumnRange(null!, out _, out _, out _))
            .Throws<ArgumentNullException>();
    }

    // === GetRegion: the six areas, three outcomes ========================================================

    // The oracle fixture, mirrored as a Table: Data!Tabela1 = A1:C4, header Item/Valor/Qtd, data rows
    // a,10,1 / b,20,2 / c,30,3; with `totals` it is A1:C5, A5 = "Total", B5 empty, C5 = SUBTOTAL(109,[Qtd]).
    private static Table Oracle(bool totals) =>
        new(
            "Tabela1",
            "Data",
            FirstRow: 1,
            LastRow: totals ? 5 : 4,
            FirstColumn: 1,
            HasHeaderRow: true,
            HasTotalsRow: totals,
            ["Item", "Valor", "Qtd"]
        );

    // Measured on Aspose.Cells 26.6.0 (2026-09-11) as ROW/ROWS/COLUMN/COLUMNS of each specifier typed on
    // Main!H20, PLAIN and array-entered — the two modes agreed on every row below. The rectangle in each
    // row is (COLUMN, ROW, COLUMN+COLUMNS-1, ROW+ROWS-1):
    //
    // | specifier              | no totals row       | with a totals row   |
    // | ---------------------- | ------------------- | ------------------- |
    // | [#All]                 | A1:C4 (ROWS 4)      | A1:C5 (ROWS 5)      |
    // | [#Data]                | A2:C4 (ROWS 3)      | A2:C4 (ROWS 3)      |
    // | [#Headers]             | A1:C1 (ROWS 1)      | A1:C1 (ROWS 1)      |
    // | [#Totals]              | #REF!               | A5:C5 (ROWS 1)      |
    // | [[#Headers],[#Data]]   | A1:C4 (ROWS 4)      | A1:C4 (ROWS 4)      |
    // | [[#Data],[#Totals]]    | A2:C4 (ROWS 3)      | A2:C5 (ROWS 4)      |
    //
    // The two PAIRS therefore SHRINK when the row they name is absent — only the singletons error. Copying
    // the singleton's error arm into a pair by analogy would answer #REF! where the oracle answers the data.
    [Test]
    [Arguments(TableArea.All, false, 1, 1, 3, 4)]
    [Arguments(TableArea.All, true, 1, 1, 3, 5)]
    [Arguments(TableArea.Data, false, 1, 2, 3, 4)]
    [Arguments(TableArea.Data, true, 1, 2, 3, 4)]
    [Arguments(TableArea.Headers, false, 1, 1, 3, 1)]
    [Arguments(TableArea.Headers, true, 1, 1, 3, 1)]
    [Arguments(TableArea.Totals, true, 1, 5, 3, 5)]
    [Arguments(TableArea.HeadersAndData, false, 1, 1, 3, 4)]
    [Arguments(TableArea.HeadersAndData, true, 1, 1, 3, 4)]
    [Arguments(TableArea.DataAndTotals, false, 1, 2, 3, 4)]
    [Arguments(TableArea.DataAndTotals, true, 1, 2, 3, 5)]
    public async Task GetRegion_WholeBand_IsTheMeasuredRectangle(
        TableArea area,
        bool totals,
        int left,
        int top,
        int right,
        int bottom
    )
    {
        var outcome = Oracle(totals)
            .GetRegion(null, area, out var l, out var t, out var r, out var b);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Resolved);
        await Assert.That((l, t, r, b)).IsEqualTo((left, top, right, bottom));
    }

    // The column narrowing happens AFTER the row band, so every area keeps its rows and gives up its
    // columns. Measured the same way: [[#All],[Valor]] = B1:B4 (ROW 1, ROWS 4, COLUMN 2, COLUMNS 1),
    // [[#Headers],[Valor]] = B1:B1, [Valor] = B2:B4, [[#Totals],[Valor]] = B5:B5 with a totals row,
    // [[#Headers],[#Data],[Valor]] = B1:B4, [[#Data],[#Totals],[Valor]] = B2:B5 with one.
    [Test]
    [Arguments(TableArea.All, false, 2, 1, 2, 4)]
    [Arguments(TableArea.All, true, 2, 1, 2, 5)]
    [Arguments(TableArea.Data, false, 2, 2, 2, 4)]
    [Arguments(TableArea.Headers, false, 2, 1, 2, 1)]
    [Arguments(TableArea.Totals, true, 2, 5, 2, 5)]
    [Arguments(TableArea.HeadersAndData, false, 2, 1, 2, 4)]
    [Arguments(TableArea.HeadersAndData, true, 2, 1, 2, 4)]
    [Arguments(TableArea.DataAndTotals, false, 2, 2, 2, 4)]
    [Arguments(TableArea.DataAndTotals, true, 2, 2, 2, 5)]
    public async Task GetRegion_NarrowedToAColumn_KeepsTheBandAndTakesOneColumn(
        TableArea area,
        bool totals,
        int left,
        int top,
        int right,
        int bottom
    )
    {
        // "valor" ignoring case, as Excel resolves Table1[col] to Table1[Col].
        var outcome = Oracle(totals)
            .GetRegion("valor", area, out var l, out var t, out var r, out var b);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Resolved);
        await Assert.That((l, t, r, b)).IsEqualTo((left, top, right, bottom));
    }

    // Only the SINGLETONS are absent when the row they name is missing. Two fixtures, kept apart because
    // their numbers are: on the A1:C4 fixture with no totals row, [#Totals] is #REF! through
    // ROW/ROWS/COLUMN/COLUMNS/SUM; on a SEPARATE header-less fixture (Aspose ShowHeaderRow = false, which
    // leaves two data rows at A2:C3 summing 55), [#Headers] is #REF! while [[#Headers],[#Data]] is that
    // fixture's data body — ROW 2, ROWS 2, SUM 55 — the shrink rule's other direction. The table below is
    // this test's own 1..4 shape and neither of those numbers describes it; what it pins is the outcome.
    [Test]
    [Arguments(TableArea.Totals, true, false)]
    [Arguments(TableArea.Headers, false, true)]
    public async Task GetRegion_AnAbsentSingletonRow_IsAbsent(
        TableArea area,
        bool hasHeaderRow,
        bool hasTotalsRow
    )
    {
        var table = new Table(
            "Tabela1",
            "Data",
            1,
            4,
            1,
            hasHeaderRow,
            hasTotalsRow,
            ["Item", "Valor", "Qtd"]
        );

        var outcome = table.GetRegion(null, area, out var l, out var t, out var r, out var b);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Absent);
        await Assert.That((l, t, r, b)).IsEqualTo((0, 0, 0, 0));
    }

    // A header-LESS table's [[#Headers],[#Data]] shrinks to the data body instead of erroring (oracle row
    // above), and a totals-LESS table's [[#Data],[#Totals]] does the same (Phase 4 ruling 2).
    [Test]
    public async Task GetRegion_APairMissingItsNamedRow_ShrinksToTheDataBody()
    {
        var headerless = new Table(
            "Tabela1",
            "Data",
            1,
            4,
            1,
            HasHeaderRow: false,
            HasTotalsRow: false,
            ["Item", "Valor", "Qtd"]
        );

        var pair = headerless.GetRegion(
            null,
            TableArea.HeadersAndData,
            out var l,
            out var t,
            out var r,
            out var b
        );
        var data = headerless.GetRegion(
            null,
            TableArea.Data,
            out var dl,
            out var dt,
            out var dr,
            out var db
        );

        await Assert.That(pair).IsEqualTo(TableRegionOutcome.Resolved);
        await Assert.That(data).IsEqualTo(TableRegionOutcome.Resolved);
        await Assert.That((l, t, r, b)).IsEqualTo((dl, dt, dr, db));
        await Assert.That((l, t, r, b)).IsEqualTo((1, 1, 3, 4));
    }

    // An unknown column is ABSENT for every area, never Empty: an absent region is #REF! while an empty one is
    // an empty reference (sweep item 33), so folding a typo'd column into the empty-body outcome would
    // silently answer SUM 0 for a column that does not exist.
    [Test]
    [Arguments(TableArea.All)]
    [Arguments(TableArea.Data)]
    [Arguments(TableArea.Headers)]
    [Arguments(TableArea.HeadersAndData)]
    [Arguments(TableArea.DataAndTotals)]
    public async Task GetRegion_UnknownColumn_IsAbsent(TableArea area)
    {
        var outcome = Oracle(false)
            .GetRegion("Imposto", area, out var l, out var t, out var r, out var b);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Absent);
        await Assert.That((l, t, r, b)).IsEqualTo((0, 0, 0, 0));
    }

    // An area value off the wire that no member names (a hand-corrupted file: Validate does not run on
    // deserialization) is Absent, not a throw and not a silent rectangle.
    [Test]
    public async Task GetRegion_AnUnknownAreaValue_IsAbsent()
    {
        var outcome = Oracle(false)
            .GetRegion(null, (TableArea)9, out var l, out var t, out var r, out var b);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Absent);
        await Assert.That((l, t, r, b)).IsEqualTo((0, 0, 0, 0));
    }

    // The THIRD outcome, kept distinct from Absent: a header-only table's data band spans zero rows. The
    // oracle calls this an EMPTY reference (SUM 0, ROWS 0, ISREF TRUE), and since sweep item 33 the outcome
    // carries the band's geometry — the columns, the top row after the header, and a bottom row above it —
    // which TableReference turns into an EmptyRangeReference. Before, the out parameters were (0, 0, 0, 0)
    // and the resolver answered #REF!. It must never become a RangeReference: measured in-tree,
    // new RangeReference("B2", "B1") reports TopRow 1, ROWS 2 and SUM 5, so the inverted pair would silently
    // read the HEADER row (AnInvertedRangeReference_IsNormalized_WhichIsWhyEmptyExists).
    [Test]
    [Arguments(TableArea.Data, null)]
    [Arguments(TableArea.Data, "Valor")]
    [Arguments(TableArea.DataAndTotals, null)]
    [Arguments(TableArea.DataAndTotals, "Valor")]
    public async Task GetRegion_AZeroRowBand_IsEmpty_WithItsGeometry(TableArea area, string? column)
    {
        // A1:C1, header row only: FirstDataRow 2 is past LastDataRow 1.
        var headerOnly = new Table(
            "Vazia",
            "Data",
            1,
            1,
            1,
            HasHeaderRow: true,
            HasTotalsRow: false,
            ["Item", "Valor", "Qtd"]
        );

        var outcome = headerOnly.GetRegion(
            column,
            area,
            out var l,
            out var t,
            out var r,
            out var b
        );

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Empty);
        await Assert.That((l, t, r, b)).IsEqualTo(column is null ? (1, 2, 3, 1) : (2, 2, 2, 1));
    }

    // The mirror shape, measured on the oracle rather than derived (Aspose: header-only Add then
    // ShowTotals, ref A1:C2 — a header row, a totals row and NOTHING between them). Only [#Data] is empty
    // there (ROWS 0); the two pairs shrink to the single row each one still HAS, which is the shrink rule
    // taken to its limit and NOT an empty band:
    //
    // | specifier            | ROW | ROWS | COUNTA | rectangle           |
    // | [#All]               | 1   | 2    | 5      | A1:C2               |
    // | [#Data]              | 2   | 0    | —      | empty               |
    // | [#Headers]           | 1   | 1    | 3      | A1:C1               |
    // | [#Totals]            | 2   | 1    | 2      | A2:C2               |
    // | [[#Headers],[#Data]] | 1   | 1    | 3      | A1:C1 (header only) |
    // | [[#Data],[#Totals]]  | 2   | 1    | 2      | A2:C2 (totals only) |
    [Test]
    [Arguments(TableArea.All, 1, 2)]
    [Arguments(TableArea.Headers, 1, 1)]
    [Arguments(TableArea.Totals, 2, 2)]
    [Arguments(TableArea.HeadersAndData, 1, 1)]
    [Arguments(TableArea.DataAndTotals, 2, 2)]
    public async Task GetRegion_AHeaderAndTotalsRowWithNoData_ShrinksToTheRowsThatExist(
        TableArea area,
        int top,
        int bottom
    )
    {
        var noData = new Table(
            "Vazia",
            "Data",
            1,
            2,
            1,
            HasHeaderRow: true,
            HasTotalsRow: true,
            ["Item", "Valor", "Qtd"]
        );

        var outcome = noData.GetRegion(null, area, out var l, out var t, out var r, out var b);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Resolved);
        await Assert.That((l, t, r, b)).IsEqualTo((1, top, 3, bottom));
    }

    // …and the one band that IS empty on that shape.
    [Test]
    public async Task GetRegion_AHeaderAndTotalsRowWithNoData_HasAnEmptyDataBand()
    {
        var noData = new Table(
            "Vazia",
            "Data",
            1,
            2,
            1,
            HasHeaderRow: true,
            HasTotalsRow: true,
            ["Item", "Valor", "Qtd"]
        );

        await Assert
            .That(noData.GetRegion(null, TableArea.Data, out _, out _, out _, out _))
            .IsEqualTo(TableRegionOutcome.Empty);
    }

    // HeadersAndData is the THIRD band whose rows can invert, and the only shape that inverts it is a table
    // with a totals row, no header row and no data — ref="A1:A1" totalsRowCount="1", which Validate accepts.
    // Not measurable on the oracle: Aspose cannot build a table with no data rows AND no header row (hiding
    // the header row consumes one row of data), so this is MySheet's own answer from the one bottom < top
    // rule, not an oracle row. Without it, the branch that stops HeadersAndData from emitting the inverted
    // (1, 0) pair is untested.
    [Test]
    public async Task GetRegion_ATotalsOnlyTable_HasAnEmptyHeadersAndDataBand()
    {
        var totalsOnly = new Table(
            "SoTotais",
            "Data",
            1,
            1,
            1,
            HasHeaderRow: false,
            HasTotalsRow: true,
            ["Item", "Valor", "Qtd"]
        );

        await Assert
            .That(totalsOnly.GetRegion(null, TableArea.HeadersAndData, out _, out _, out _, out _))
            .IsEqualTo(TableRegionOutcome.Empty);
        await Assert
            .That(totalsOnly.GetRegion(null, TableArea.Data, out _, out _, out _, out _))
            .IsEqualTo(TableRegionOutcome.Empty);
        await Assert
            .That(totalsOnly.GetRegion(null, TableArea.Headers, out _, out _, out _, out _))
            .IsEqualTo(TableRegionOutcome.Absent);

        // The bands that still exist are the single row the table has.
        await Assert
            .That(
                totalsOnly.GetRegion(
                    null,
                    TableArea.DataAndTotals,
                    out _,
                    out var pairTop,
                    out _,
                    out var pairBottom
                )
            )
            .IsEqualTo(TableRegionOutcome.Resolved);
        await Assert.That((pairTop, pairBottom)).IsEqualTo((1, 1));
        await Assert
            .That(totalsOnly.GetRegion(null, TableArea.All, out _, out _, out _, out _))
            .IsEqualTo(TableRegionOutcome.Resolved);
        await Assert
            .That(totalsOnly.GetRegion(null, TableArea.Totals, out _, out _, out _, out _))
            .IsEqualTo(TableRegionOutcome.Resolved);
    }

    // An unknown column on a table with no data rows reports the COLUMN's absence, not the empty band: the
    // column check runs first precisely so a typo never hides behind a data-dependent outcome.
    [Test]
    public async Task GetRegion_UnknownColumnOnAnEmptyTable_IsAbsentNotEmpty()
    {
        var headerOnly = new Table(
            "Vazia",
            "Data",
            1,
            1,
            1,
            HasHeaderRow: true,
            HasTotalsRow: false,
            ["Item", "Valor", "Qtd"]
        );

        var outcome = headerOnly.GetRegion("Imposto", TableArea.Data, out _, out _, out _, out _);

        await Assert.That(outcome).IsEqualTo(TableRegionOutcome.Absent);
    }

    // A ref that cannot be a rectangle is Absent for EVERY area, and reaches this method only through
    // deserialization (Validate rejects both shapes). Before the guard, a zero-column table answered
    // Resolved with left 1 and right 0, TryResolveRange built CellAddress(0, 4).ToId() == "4" from it, and
    // SUM / ROWS / COLUMNS / Evaluate all THREW FormatException "Invalid cell reference '4'" — measured, and
    // a throw is the one thing TableReference.Evaluate's contract rules out.
    [Test]
    [Arguments(TableArea.All)]
    [Arguments(TableArea.Data)]
    [Arguments(TableArea.Headers)]
    [Arguments(TableArea.Totals)]
    [Arguments(TableArea.HeadersAndData)]
    [Arguments(TableArea.DataAndTotals)]
    public async Task GetRegion_AnImpossibleRef_IsAbsentForEveryArea(TableArea area)
    {
        // No columns at all: LastColumn (0) is before FirstColumn (1).
        var noColumns = new Table("Corrupt", "Data", 1, 4, 1, true, true, []);
        // LastRow before FirstRow: every band would be inverted or outside the ref.
        var backwards = new Table("Corrupt", "Data", 5, 3, 1, true, true, ["Item", "Valor"]);

        await Assert.That(noColumns.LastColumn).IsEqualTo(0);
        await Assert
            .That(noColumns.GetRegion(null, area, out var l, out var t, out var r, out var b))
            .IsEqualTo(TableRegionOutcome.Absent);
        await Assert.That((l, t, r, b)).IsEqualTo((0, 0, 0, 0));
        await Assert
            .That(backwards.GetRegion(null, area, out var bl, out var bt, out var br, out var bb))
            .IsEqualTo(TableRegionOutcome.Absent);
        await Assert.That((bl, bt, br, bb)).IsEqualTo((0, 0, 0, 0));
    }

    // The measurement the Empty outcome exists for, pinned where it can rot: an inverted pair is NOT a safe
    // way to say "no rows". RangeReference normalizes min/max, so (B2, B1) is a real two-cell rectangle
    // starting at row 1 — which for a table means the HEADER row — and SUM over it reads both cells. If this
    // ever goes red because GetBounds started rejecting an inverted pair, TableRegionOutcome.Empty's
    // rationale has changed and its doc must change with it.
    [Test]
    public async Task AnInvertedRangeReference_IsNormalized_WhichIsWhyEmptyExists()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["B1"] = new NumberValue(2);
        data["B2"] = new NumberValue(3);
        var inverted = new RangeReference("B2", "B1", "Data");

        await Assert.That(inverted.TopRow).IsEqualTo(1);
        await Assert.That(inverted.RowCount).IsEqualTo(2);
        await Assert
            .That(
                new Sum([inverted]).Evaluate(new EvaluationContext(workbook)).AsObject() as double?
            )
            .IsEqualTo(5.0);
    }
}
