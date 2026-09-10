namespace Danfma.MySheet.Tests;

/// <summary>
/// The <see cref="Table"/> record on its own — the geometry it derives from the xlsx-shaped
/// <c>FirstRow/LastRow/FirstColumn + HasHeaderRow/HasTotalsRow</c> members, the column lookup, and the
/// [#Data] range it hands the reference-semantics phase. Nothing here touches a <see cref="Workbook"/>: the
/// registry's pins live in <c>TableRegistryTests</c>.
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
}
