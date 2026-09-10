namespace Danfma.MySheet.Tests;

#if MYSHEET_TABLES
/// <summary>
/// Acceptance pins for the table registry: <see cref="Table"/>'s geometry and column lookup,
/// <see cref="Workbook.Tables"/>/<c>DefineTable</c>, the table-name rule, the shared name/table namespace and
/// the registry's survival across Save/Load. Gated behind <c>MYSHEET_TABLES</c> (see the test .csproj) until the
/// model lands; these are the acceptance criteria for the task that lands it.
/// </summary>
public class TableRegistryTests
{
    // A well-formed two-column table over Data!A1:B4 — header row 1, data rows 2..4, no totals row.
    private static Table Sample(string name = "Vendas") =>
        new(name, "Data", 1, 4, 1, HasHeaderRow: true, HasTotalsRow: false, ["Produto", "Total"]);

    // === The registry ====================================================================================

    [Test]
    public async Task DefineTable_ThenLookup_IsCaseInsensitive()
    {
        var wb = new Workbook();
        wb.DefineTable(Sample());

        await Assert.That(wb.Tables.Count).IsEqualTo(1);
        await Assert.That(wb.Tables.ContainsKey("vendas")).IsTrue();
        await Assert.That(wb.Tables["VENDAS"].Name).IsEqualTo("Vendas");
    }

    // Redefining REPLACES: the registry is a map, not an append log. The key keeps its original casing (the
    // dictionary is OrdinalIgnoreCase), so only the geometry is asserted here.
    [Test]
    public async Task Redefinition_Replaces_AndKeepsOneEntry()
    {
        var wb = new Workbook();
        wb.DefineTable(Sample());
        wb.DefineTable(new Table("vendas", "Data", 10, 20, 3, true, true, ["A", "B", "C"]));

        var table = wb.Tables["Vendas"];

        await Assert.That(wb.Tables.Count).IsEqualTo(1);
        await Assert.That(table.FirstRow).IsEqualTo(10);
        await Assert.That(table.LastRow).IsEqualTo(20);
        await Assert.That(table.FirstColumn).IsEqualTo(3);
        await Assert.That(table.ColumnCount).IsEqualTo(3);
        await Assert.That(table.HasTotalsRow).IsTrue();
    }

    // === Derived geometry ================================================================================

    // FirstRow/LastRow span the WHOLE xlsx ref, totals row included: ref="A1:B4" totalsRowCount="1" means
    // header 1, data 2..3, totals 4 (measured in a real ClosedXML-authored file).
    [Test]
    public async Task A1Overload_WithTotalsRow_DerivesTheDataRows()
    {
        var wb = new Workbook();
        wb.DefineTable("Vendas", "Data", "A1:B4", ["Produto", "Total"], hasTotalsRow: true);

        var table = wb.Tables["Vendas"];

        await Assert.That(table.FirstRow).IsEqualTo(1);
        await Assert.That(table.LastRow).IsEqualTo(4);
        await Assert.That(table.FirstColumn).IsEqualTo(1);
        await Assert.That(table.HeaderRow).IsEqualTo(1);
        await Assert.That(table.FirstDataRow).IsEqualTo(2);
        await Assert.That(table.LastDataRow).IsEqualTo(3);
        await Assert.That(table.DataRowCount).IsEqualTo(2);
        await Assert.That(table.TotalsRow).IsEqualTo(4);
        await Assert.That(table.LastColumn).IsEqualTo(2);
    }

    // A header row and nothing else is LEGAL, and the zero-data-row case falls out of the arithmetic instead of
    // needing a sentinel: [#Data] has no rows, so TryGetColumnRange has no range to hand back.
    [Test]
    public async Task HeaderOnlyTable_HasNoDataRows_AndNoColumnRange()
    {
        var wb = new Workbook();
        wb.DefineTable("Solo", "Data", "A1:A1", ["Produto"]);

        var table = wb.Tables["Solo"];

        await Assert.That(table.DataRowCount).IsEqualTo(0);
        await Assert.That(table.TryGetColumnRange("Produto", out _, out _, out _)).IsFalse();
    }

    [Test]
    public async Task TryGetColumnIndex_IsCaseInsensitive_AndFalseForUnknown()
    {
        var table = Sample();

        await Assert.That(table.TryGetColumnIndex("PRODUTO", out var produto)).IsTrue();
        await Assert.That(produto).IsEqualTo(0);
        await Assert.That(table.TryGetColumnIndex("total", out var total)).IsTrue();
        await Assert.That(total).IsEqualTo(1);
        await Assert.That(table.TryGetColumnIndex("Imposto", out _)).IsFalse();
    }

    // === Validation ======================================================================================

    [Test]
    public async Task DuplicateColumnName_DifferingOnlyInCase_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() =>
                wb.DefineTable(new Table("T", "Data", 1, 4, 1, true, false, ["Total", "total"]))
            )
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task EmptyColumnName_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() =>
                wb.DefineTable(new Table("T", "Data", 1, 4, 1, true, false, ["Total", "  "]))
            )
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task ZeroColumns_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable(new Table("T", "Data", 1, 4, 1, true, false, [])))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task EmptySheetName_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable(new Table("T", "  ", 1, 4, 1, true, false, ["Total"])))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task LastRowBeforeFirstRow_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable(new Table("T", "Data", 5, 4, 1, true, false, ["Total"])))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task RefWidthNotMatchingColumnCount_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable("T", "Data", "A1:B4", ["Produto"]))
            .Throws<ArgumentException>();
    }

    // === The A1 overload's `reference` parsing ============================================================
    //
    // `reference` is the BARE xlsx <table ref="…"> string, so the only two accepted shapes are one A1 corner
    // ("C2") and two separated by a single ':'. These cases were implemented alongside the overload itself
    // (Workbook.TryParseTableReference), so every assertion below is GREEN on arrival: they are regression
    // pins for the malformed-input rejections, not a red-to-green cycle.

    [Test]
    [Arguments("A1:B4:C5")] // two colons: the text after the first ':' is not a bare A1 corner
    [Arguments("Data!A1:B4")] // a sheet qualifier: `sheetName` carries the sheet, `reference` never does
    [Arguments("A0:B4")] // row 0 does not exist (CellAddress.TryParseRow rejects it)
    public async Task A1Overload_MalformedReference_Throws(string reference)
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable("T", "Data", reference, ["Produto", "Total"]))
            .Throws<ArgumentException>();
    }

    // The two shapes that DO parse: CellAddress.TryParseA1 skips the absolute markers and
    // TryParseTableReference normalizes the corners with Math.Min/Math.Max, so both register the same geometry
    // as the plain "A1:B4" of Sample().
    [Test]
    [Arguments("$A$1:$B$4")]
    [Arguments("B4:A1")] // reversed corners
    public async Task A1Overload_AbsoluteMarkersAndReversedCorners_AreAccepted(string reference)
    {
        var wb = new Workbook();

        wb.DefineTable("T", "Data", reference, ["Produto", "Total"]);
        var table = wb.Tables["T"];

        await Assert.That(table.FirstRow).IsEqualTo(1);
        await Assert.That(table.LastRow).IsEqualTo(4);
        await Assert.That(table.FirstColumn).IsEqualTo(1);
        await Assert.That(table.LastColumn).IsEqualTo(2);
    }

    [Test]
    public async Task FirstRowZero_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable(new Table("T", "Data", 0, 4, 1, true, false, ["Total"])))
            .Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task FirstColumnZero_Throws()
    {
        var wb = new Workbook();

        await Assert
            .That(() => wb.DefineTable(new Table("T", "Data", 1, 4, 0, true, false, ["Total"])))
            .Throws<ArgumentOutOfRangeException>();
    }

    // === The shared name/table namespace =================================================================

    [Test]
    public async Task NameCollidingWithADefinedName_Throws()
    {
        var wb = new Workbook();
        wb.Sheets.Add("Data");
        wb.DefineName("Vendas", "Data!A1:A3");

        await Assert.That(() => wb.DefineTable(Sample())).Throws<ArgumentException>();
    }

    // The guard is symmetric: a defined name may not shadow a registered table either.
    [Test]
    public async Task DefinedNameCollidingWithATable_Throws()
    {
        var wb = new Workbook();
        wb.Sheets.Add("Data");
        wb.DefineTable(Sample());

        await Assert.That(() => wb.DefineName("Vendas", "Data!A1:A3")).Throws<ArgumentException>();
    }

    // === The table-name rule =============================================================================

    // "Tabela1" and "Table1" are Excel's OWN default table names (measured in a real ClosedXML-authored file:
    // name="Tabela1" displayName="Tabela1") and they are exactly what NamedReferences.ValidateName REJECTS
    // today, because Parser.IsCellReference treats any letters-then-digits string as a cell reference with no
    // column/row bound. This assertion is the regression guard for the dedicated, grid-bounded check.
    [Test]
    public async Task ExcelDefaultAndDottedNames_AreAccepted()
    {
        foreach (var name in new[] { "Tabela1", "Table1", "Vendas.2024", "_x" })
        {
            var wb = new Workbook();
            wb.DefineTable(Sample(name));

            await Assert.That(wb.Tables.ContainsKey(name)).IsTrue();
        }
    }

    // "A1" and "XFD1048576" are inside Excel's grid, so they are ambiguous with a cell reference; "R1C1" is the
    // R1C1 shape; "C" is one of the four reserved single letters; the last three break the character rule or the
    // 255-character limit.
    [Test]
    public async Task GridReferencesReservedAndMalformedNames_AreRejected()
    {
        string[] rejected =
        [
            "A1",
            "XFD1048576",
            "R1C1",
            "C",
            "My Table",
            "Table-1",
            new string('T', 256),
        ];

        foreach (var name in rejected)
        {
            var wb = new Workbook();

            await Assert.That(() => wb.DefineTable(Sample(name))).Throws<ArgumentException>();
        }
    }

    // === Persistence =====================================================================================

    [Test]
    public async Task SaveLoad_PreservesEveryMember()
    {
        var wb = new Workbook();
        wb.Sheets.Add("Data");
        wb.DefineTable(
            "Tabela1",
            "Data",
            "C2:E9",
            ["Produto", "Qtd", "Total"],
            hasHeaderRow: true,
            hasTotalsRow: true
        );

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            wb.Save(path);
            var loaded = Workbook.Load(path);
            var table = loaded.Tables["Tabela1"];

            await Assert.That(loaded.Tables.Count).IsEqualTo(1);
            // The OrdinalIgnoreCase comparer is restored after deserialization, not merely on the fresh map.
            await Assert.That(loaded.Tables.ContainsKey("tabela1")).IsTrue();
            await Assert.That(table.Name).IsEqualTo("Tabela1");
            await Assert.That(table.SheetName).IsEqualTo("Data");
            await Assert.That(table.FirstRow).IsEqualTo(2);
            await Assert.That(table.LastRow).IsEqualTo(9);
            await Assert.That(table.FirstColumn).IsEqualTo(3);
            await Assert.That(table.HasHeaderRow).IsTrue();
            await Assert.That(table.HasTotalsRow).IsTrue();
            await Assert.That(table.ColumnNames).IsEquivalentTo(["Produto", "Qtd", "Total"]);
            await Assert.That(table.LastColumn).IsEqualTo(5);

            // The lazy column index rebuilds over the List-backed ColumnNames a Load produces.
            await Assert.That(table.TryGetColumnIndex("QTD", out var qtd)).IsTrue();
            await Assert.That(qtd).IsEqualTo(1);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === The folded definitions counter ==================================================================

    // One counter for BOTH namespaces: RecalculationEngine.EnsureFresh only ever rebuilds the whole graph, so
    // splitting names from tables would buy nothing.
    [Test]
    public async Task DefinitionsVersion_AdvancesOnEachDefineTable()
    {
        var wb = new Workbook();
        wb.Sheets.Add("Data");

        var start = wb.DefinitionsVersion;
        wb.DefineTable(Sample("Vendas"));
        var afterFirst = wb.DefinitionsVersion;
        wb.DefineTable(Sample("Custos"));
        var afterSecond = wb.DefinitionsVersion;
        wb.DefineName("N", "Data!A1");
        var afterName = wb.DefinitionsVersion;

        await Assert.That(afterFirst).IsEqualTo(start + 1);
        await Assert.That(afterSecond).IsEqualTo(start + 2);
        await Assert.That(afterName).IsEqualTo(start + 3);
    }
}
#endif
