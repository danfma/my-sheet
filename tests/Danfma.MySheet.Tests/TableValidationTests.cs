namespace Danfma.MySheet.Tests;

/// <summary>
/// <see cref="Table.ValidateName"/> (Excel's documented table-name rule, plus the MySheet-specific
/// TRUE/FALSE and backslash rejections) and <see cref="Table.Validate"/> (the structural checks the registry
/// runs at the registration boundary — never in the constructor, so a loaded file cannot throw during
/// deserialization).
/// </summary>
public class TableValidationTests
{
    private static Table Valid(string name = "Vendas") =>
        new(name, "Data", 1, 4, 1, HasHeaderRow: true, HasTotalsRow: false, ["Produto", "Total"]);

    // === ValidateName ====================================================================================

    [Test]
    [Arguments("Tabela1")] // Excel's own default names: letters-then-digits, but past the 3-letter bound
    [Arguments("Table1")]
    [Arguments("Vendas.2024")]
    [Arguments("Sales.Data")]
    [Arguments("_x")]
    [Arguments("T")] // a single letter other than C/c/R/r
    [Arguments("XFE1")] // letters-then-digits OUTSIDE the grid: not a cell, so a legal name
    [Arguments("A1048577")]
    [Arguments("R1C")] // not the full R<digits>C<digits> shape
    [Arguments("Ação")] // non-ASCII letters are letters
    public async Task ValidateName_AcceptsLegalNames(string name)
    {
        await Assert.That(() => Table.ValidateName(name)).ThrowsNothing();
    }

    [Test]
    public async Task ValidateName_Accepts255Characters_AndRejects256()
    {
        await Assert.That(() => Table.ValidateName(new string('T', 255))).ThrowsNothing();
        await Assert
            .That(() => Table.ValidateName(new string('T', 256)))
            .Throws<ArgumentException>();
    }

    [Test]
    [Arguments("")]
    [Arguments("   ")]
    [Arguments("1abc")] // must start with a letter or '_'
    [Arguments(".abc")]
    [Arguments("My Table")] // no spaces
    [Arguments("Table-1")] // only letters, digits, '.' and '_'
    [Arguments("Tab\\1")] // Excel allows a leading backslash; the tokenizer could never read it
    [Arguments("C")] // the four reserved single letters
    [Arguments("c")]
    [Arguments("R")]
    [Arguments("r")]
    [Arguments("A1")] // inside the grid
    [Arguments("T1")]
    [Arguments("XFD1048576")]
    [Arguments("R1C1")] // the R1C1 shape, either case
    [Arguments("r2c3")]
    [Arguments("TRUE")] // the tokenizer reads these as BooleanValue, so the name could never be reached
    [Arguments("false")]
    public async Task ValidateName_RejectsIllegalNames(string name)
    {
        await Assert.That(() => Table.ValidateName(name)).Throws<ArgumentException>();
    }

    [Test]
    public async Task ValidateName_QuotesTheOffendingValue_AndNamesTheParameter()
    {
        var exception = await Assert
            .That(() => Table.ValidateName("My Table"))
            .Throws<ArgumentException>()
            .WithParameterName("name");

        await Assert.That(exception).IsNotNull();
        await Assert.That(exception!.Message).Contains("'My Table'");
    }

    [Test]
    public async Task ValidateName_Null_Throws()
    {
        await Assert.That(() => Table.ValidateName(null!)).Throws<ArgumentNullException>();
    }

    // === Validate ========================================================================================

    [Test]
    public async Task Validate_WellFormedTable_ThrowsNothing()
    {
        await Assert.That(() => Valid().Validate()).ThrowsNothing();
    }

    [Test]
    public async Task Validate_RunsTheNameRule()
    {
        await Assert.That(() => Valid("A1").Validate()).Throws<ArgumentException>();
    }

    [Test]
    public async Task Validate_SheetName_NullOrBlank_Throws()
    {
        await Assert
            .That(() => (Valid() with { SheetName = null! }).Validate())
            .Throws<ArgumentNullException>();
        await Assert
            .That(() => (Valid() with { SheetName = "  " }).Validate())
            .Throws<ArgumentException>();
    }

    // Geometry is 1-based, and the exception carries the offending value like CellRef.Format's does.
    [Test]
    public async Task Validate_FirstRowBelowOne_ThrowsOutOfRange_NamingFirstRow()
    {
        var exception = await Assert
            .That(() => (Valid() with { FirstRow = 0 }).Validate())
            .Throws<ArgumentOutOfRangeException>()
            .WithParameterName("FirstRow");

        await Assert.That(exception).IsNotNull();
        await Assert.That(exception!.ActualValue).IsEqualTo(0);
    }

    [Test]
    public async Task Validate_FirstColumnBelowOne_ThrowsOutOfRange_NamingFirstColumn()
    {
        var exception = await Assert
            .That(() => (Valid() with { FirstColumn = -3 }).Validate())
            .Throws<ArgumentOutOfRangeException>()
            .WithParameterName("FirstColumn");

        await Assert.That(exception).IsNotNull();
        await Assert.That(exception!.ActualValue).IsEqualTo(-3);
    }

    [Test]
    public async Task Validate_LastRowBeforeFirstRow_Throws()
    {
        await Assert
            .That(() => (Valid() with { FirstRow = 5, LastRow = 4 }).Validate())
            .Throws<ArgumentException>();
    }

    // A one-row table cannot hold BOTH a header and a totals row; the message names the two flags.
    [Test]
    public async Task Validate_SpanShorterThanHeaderPlusTotals_Throws_NamingBothFlags()
    {
        var table = new Table("T", "Data", 1, 1, 1, HasHeaderRow: true, HasTotalsRow: true, ["A"]);

        var exception = await Assert
            .That(() => table.Validate())
            .Throws<ArgumentException>()
            .WithMessageContaining("HasHeaderRow");

        await Assert.That(exception).IsNotNull();
        await Assert.That(exception!.Message).Contains("HasTotalsRow");
    }

    // Zero data rows is EXPLICITLY legal: header only, totals only, or header + totals over two rows.
    [Test]
    [Arguments(1, 1, true, false)]
    [Arguments(1, 1, false, true)]
    [Arguments(1, 2, true, true)]
    [Arguments(1, 1, false, false)] // one data row and nothing else
    public async Task Validate_SpanEqualToHeaderPlusTotals_IsLegal(
        int firstRow,
        int lastRow,
        bool hasHeaderRow,
        bool hasTotalsRow
    )
    {
        var table = new Table("T", "Data", firstRow, lastRow, 1, hasHeaderRow, hasTotalsRow, ["A"]);

        await Assert.That(() => table.Validate()).ThrowsNothing();
    }

    [Test]
    public async Task Validate_NoColumns_Throws()
    {
        await Assert
            .That(() => (Valid() with { ColumnNames = [] }).Validate())
            .Throws<ArgumentException>();
    }

    [Test]
    [Arguments(null)]
    [Arguments("")]
    [Arguments("  ")]
    public async Task Validate_BlankColumnName_Throws_NamingItsIndex(string? blank)
    {
        var table = Valid() with { ColumnNames = ["Produto", blank!, "Total"] };

        await Assert
            .That(() => table.Validate())
            .Throws<ArgumentException>()
            .WithMessageContaining("at index 1");
    }

    // Column names are human-authored ('Valor Total (R$)' in a measured fixture) and never see the name
    // rule — spaces, parentheses and currency signs are all fine.
    [Test]
    public async Task Validate_HumanAuthoredColumnNames_AreNotSubjectToTheNameRule()
    {
        var table = Valid() with { ColumnNames = ["Valor Total (R$)", "A1", "TRUE", "C"] };

        await Assert.That(() => table.Validate()).ThrowsNothing();
    }

    // Excel resolves Table1[col] to Table1[Col], so uniqueness must be case-insensitive (ClosedXML refuses
    // 'Col'/'col' the same way). The message names both positions and the repeated text.
    [Test]
    public async Task Validate_DuplicateColumnName_IgnoringCase_Throws_NamingBothIndices()
    {
        var table = Valid() with { ColumnNames = ["Produto", "Total", "TOTAL"] };

        var exception = await Assert
            .That(() => table.Validate())
            .Throws<ArgumentException>()
            .WithMessageContaining("'TOTAL'");

        await Assert.That(exception).IsNotNull();
        await Assert.That(exception!.Message).Contains("at index 2");
        await Assert.That(exception!.Message).Contains("at index 1");
    }
}
