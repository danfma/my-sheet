using System.IO.Compression;
using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// The loader reads each worksheet's <c>&lt;table&gt;</c> parts into <see cref="Workbook.Tables"/>
/// (<c>TableDefinitionReader</c>). These are REGISTRY-level pins: which <see cref="Table"/> record lands
/// in the registry, with what geometry and column names, and which warning a malformed part raises.
/// Whether a structured reference into that table then EVALUATES is <see cref="TableInteropTests"/>'
/// business. The well-formed fixtures are oracle-authored (see Fixtures/README.md); the malformed parts
/// are ClosedXML tables whose <c>xl/tables/table1.xml</c> is overwritten afterwards, because no producer
/// writes them and the schema validator accepts most of them.
/// </summary>
public class TableDefinitionReaderTests
{
    // === Helpers ==========================================================================================

    private static (Workbook Workbook, List<ExcelLoadWarning> Warnings) Load(string path)
    {
        var warnings = new List<ExcelLoadWarning>();
        var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

        return (workbook, warnings);
    }

    private static List<ExcelLoadWarning> TableWarnings(List<ExcelLoadWarning> warnings) =>
        warnings.Where(w => w.Kind == ExcelLoadWarningKind.InvalidTableDefinition).ToList();

    // The oracle fixtures carry structured-reference formulas. Until the parser reads `[` (Phase 4) each
    // such cell degrades with UnparsableFormula; once it does, the list is empty. Both states satisfy
    // this, and neither is what these tests pin — the reader's own warning kind is.
    private static async Task AssertNoTableWarnings(List<ExcelLoadWarning> warnings)
    {
        await Assert.That(TableWarnings(warnings)).IsEmpty();
        await Assert
            .That(warnings.All(w => w.Kind == ExcelLoadWarningKind.UnparsableFormula))
            .IsTrue();
    }

    // Member-wise, because Table's synthesized equality compares ColumnNames by reference.
    private static async Task AssertTable(
        Table table,
        string name,
        string sheetName,
        int firstRow,
        int lastRow,
        int firstColumn,
        bool hasHeaderRow,
        bool hasTotalsRow,
        params string[] columnNames
    )
    {
        await Assert.That(table.Name).IsEqualTo(name);
        await Assert.That(table.SheetName).IsEqualTo(sheetName);
        await Assert.That(table.FirstRow).IsEqualTo(firstRow);
        await Assert.That(table.LastRow).IsEqualTo(lastRow);
        await Assert.That(table.FirstColumn).IsEqualTo(firstColumn);
        await Assert.That(table.HasHeaderRow).IsEqualTo(hasHeaderRow);
        await Assert.That(table.HasTotalsRow).IsEqualTo(hasTotalsRow);
        // Order matters (K|V is not V|K), so compare the joined sequence rather than the set.
        await Assert
            .That(string.Join("|", table.ColumnNames))
            .IsEqualTo(string.Join("|", columnNames));
    }

    /// <summary>A "Data" sheet holding a ClosedXML table over A1:B3 (header + two rows, B3 = 32).</summary>
    private static string WriteClosedXmlTable(
        string tableName = "Tabela1",
        Action<IXLWorksheet>? extra = null
    )
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-tablepart-{Guid.NewGuid():N}.xlsx");

        using var fixture = new XLWorkbook();
        var data = fixture.AddWorksheet("Data");

        data.Cell("A1").Value = "Item";
        data.Cell("B1").Value = "Valor";
        data.Cell("A2").Value = "a";
        data.Cell("B2").Value = 10;
        data.Cell("A3").Value = "b";
        data.Cell("B3").Value = 32;
        data.Range("A1:B3").CreateTable(tableName);
        extra?.Invoke(data);

        fixture.SaveAs(path);

        return path;
    }

    private const string PartHead =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" ";

    /// <summary>A whole <c>&lt;table&gt;</c> part from its attributes and its <c>&lt;tableColumn&gt;</c> rows.</summary>
    private static string TablePart(string attributes, string columns) =>
        $"{PartHead}{attributes}><tableColumns>{columns}</tableColumns></table>";

    /// <summary>
    /// Replaces the bytes of the one <c>xl/tables/*.xml</c> entry whose text contains
    /// <paramref name="containing"/> (every entry when null). ZipArchive rather than the SDK, because the
    /// SDK would refuse to write half of what these tests need to plant.
    /// </summary>
    private static void OverwriteTablePart(string path, string xml, string? containing = null)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip
            .Entries.Where(e => e.FullName.StartsWith("xl/tables/", StringComparison.Ordinal))
            .Single(e =>
            {
                if (containing is null)
                {
                    return true;
                }

                using var reader = new StreamReader(e.Open());
                return reader.ReadToEnd().Contains(containing, StringComparison.Ordinal);
            });
        var name = entry.FullName;

        entry.Delete();
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(xml);
    }

    // === Oracle fixtures: the record each well-formed part maps to =======================================

    [Test]
    public async Task Load_PlainTable_RegistersItsRecordAndGeometry()
    {
        var (workbook, warnings) = Load(XlsxParts.Fixture("f1-plain"));

        await AssertNoTableWarnings(warnings);
        await Assert.That(workbook.Tables.Count).IsEqualTo(1);

        var table = workbook.Tables["Tabela1"];

        await AssertTable(table, "Tabela1", "Data", 1, 3, 1, true, false, "Item", "Valor");
        await Assert.That(table.HeaderRow).IsEqualTo(1);
        await Assert.That(table.FirstDataRow).IsEqualTo(2);
        await Assert.That(table.LastDataRow).IsEqualTo(3);
        await Assert.That(table.TotalsRow).IsNull();
        await Assert.That(table.LastColumn).IsEqualTo(2);

        // The registry is case-insensitive, like Excel's name resolution.
        await Assert.That(workbook.Tables.ContainsKey("TABELA1")).IsTrue();
        // The table's cells are still ordinary cells.
        await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
    }

    [Test]
    public async Task Load_TotalsRowTable_RefIncludesTheTotalsRow()
    {
        // The part says ref="A1:B4" totalsRowCount="1" (no totalsRowShown): the ref spans the totals row,
        // so the data band must stop one row short of it or SUM(Tabela1[Valor]) double-counts the total.
        var (workbook, warnings) = Load(XlsxParts.Fixture("f2-totals"));

        await AssertNoTableWarnings(warnings);

        var table = workbook.Tables["Tabela1"];

        await AssertTable(table, "Tabela1", "Data", 1, 4, 1, true, true, "Item", "Valor");
        await Assert.That(table.TotalsRow).IsEqualTo(4);
        await Assert.That(table.FirstDataRow).IsEqualTo(2);
        await Assert.That(table.LastDataRow).IsEqualTo(3);
    }

    [Test]
    public async Task Load_HeaderlessTable_HasNoHeaderRow()
    {
        // The producer rewrote the table to ref="A2:B3" headerRowCount="0" and kept the REAL column names
        // (Item/Valor) in the part; row 1 of the sheet is empty.
        var (workbook, warnings) = Load(XlsxParts.Fixture("f3-noheader"));

        await AssertNoTableWarnings(warnings);

        var table = workbook.Tables["Tabela1"];

        await AssertTable(table, "Tabela1", "Data", 2, 3, 1, false, false, "Item", "Valor");
        await Assert.That(table.HeaderRow).IsNull();
        await Assert.That(table.FirstDataRow).IsEqualTo(2);
        await Assert.That(table.LastDataRow).IsEqualTo(3);
    }

    [Test]
    public async Task Load_HeaderOnlyTable_RegistersWithZeroDataRows()
    {
        // ref="A1:B1" with a header row: a legal table with an empty body, registered as such (the oracle
        // answers SUM(Tabela1[Valor]) = 0 and ROWS(Tabela1[#Data]) = 0 over it — Phase 5's concern).
        var (workbook, warnings) = Load(XlsxParts.Fixture("f7-header-only"));

        await AssertNoTableWarnings(warnings);

        var table = workbook.Tables["Tabela1"];

        await AssertTable(table, "Tabela1", "Data", 1, 1, 1, true, false, "Item", "Valor");
        await Assert.That(table.DataRowCount).IsEqualTo(0);
    }

    [Test]
    public async Task Load_ColumnNames_AreStoredRaw_WithOnlyTheControlEscapeDecoded()
    {
        // Spaces, parentheses and the apostrophe travel byte-for-byte (the formula-side `'` escaping is the
        // LEXER's business, never the part's); the line break the producer wrote as _x000a_ is decoded to a
        // real newline; the padded name keeps its padding.
        var (workbook, warnings) = Load(XlsxParts.Fixture("f4-names"));

        await AssertNoTableWarnings(warnings);
        await AssertTable(
            workbook.Tables["Tabela1"],
            "Tabela1",
            "Data",
            1,
            3,
            1,
            true,
            false,
            "AMOUNT IN USD For Line 1",
            "(A) NAME OF PFIC",
            "Owner's Share",
            "Line\nBreak",
            " Padded "
        );
    }

    [Test]
    public async Task Load_TwoTablesOnOneSheetAndOneOnAnother_RegistersAllThree()
    {
        var (workbook, warnings) = Load(XlsxParts.Fixture("f8-two-tables"));

        await AssertNoTableWarnings(warnings);
        await Assert.That(workbook.Tables.Count).IsEqualTo(3);
        await AssertTable(
            workbook.Tables["Tabela1"],
            "Tabela1",
            "Data",
            1,
            3,
            1,
            true,
            false,
            "Item",
            "Valor"
        );
        await AssertTable(
            workbook.Tables["Tabela2"],
            "Tabela2",
            "Data",
            1,
            3,
            6,
            true,
            false,
            "K",
            "V"
        );
        await AssertTable(
            workbook.Tables["Tabela3"],
            "Tabela3",
            "Other",
            1,
            3,
            1,
            true,
            false,
            "X"
        );
    }

    [Test]
    public async Task Load_CrossSheetFixture_RecordsTheSheetTheTableLivesOn()
    {
        // The formulas sit on Report; the table is on Data. SheetName is the table's sheet, taken from the
        // per-sheet loop that owns the worksheet part — not from any formula that mentions it.
        var (workbook, warnings) = Load(XlsxParts.Fixture("f6-cross-sheet"));

        await AssertNoTableWarnings(warnings);
        await Assert.That(workbook.Tables.Count).IsEqualTo(1);
        await Assert.That(workbook.Tables["Tabela1"].SheetName).IsEqualTo("Data");
        await Assert.That(workbook["Report"].Index).IsEqualTo(1);
    }

    [Test]
    public async Task Load_EveryOracleFixture_LeavesDefinedNamesEmpty()
    {
        // A table is never smuggled in as a defined name: the two registries are distinct.
        foreach (
            var name in new[]
            {
                "f1-plain",
                "f2-totals",
                "f3-noheader",
                "f4-names",
                "f5-name-My_Table",
                "f5-name-TRUE",
                "f6-cross-sheet",
                "f7-header-only",
                "f8-two-tables",
            }
        )
        {
            var (workbook, _) = Load(XlsxParts.Fixture(name));

            await Assert.That(workbook.DefinedNames.Count).IsEqualTo(0);
            await Assert.That(workbook.DefinedNames.ContainsKey("Tabela1")).IsFalse();
        }
    }

    [Test]
    public async Task SaveThenLoad_RoundTripsTheLoadedTable()
    {
        // The loader is the first real producer of the serialized Tables member (Phase 3 shipped it with
        // a hand-registered table only).
        var (loaded, _) = Load(XlsxParts.Fixture("f1-plain"));
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            loaded.Save(path);
            var reloaded = Workbook.Load(path);

            await Assert.That(reloaded.Tables.Count).IsEqualTo(1);
            await AssertTable(
                reloaded.Tables["Tabela1"],
                "Tabela1",
                "Data",
                1,
                3,
                1,
                true,
                false,
                "Item",
                "Valor"
            );
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Names Excel accepts and MySheet's tokenizer cannot read =========================================

    [Test]
    [Arguments("f5-name-My_Table", "My\\Table")]
    [Arguments("f5-name-TRUE", "TRUE")]
    public async Task Load_TableNameTheTokenizerCannotRead_WarnsSkips_AndKeepsTheCachedNumber(
        string fixture,
        string displayName
    )
    {
        // Excel (the oracle) accepts both names and answers 42. MySheet's tokenizer never reads `\` or
        // TRUE into an identifier, which is why Table.ValidateName rejects them: the table is skipped with
        // ONE InvalidTableDefinition naming it, the formula cell ALSO degrades (the same tokenizer), and
        // Excel's cached 42 survives in the cell — a structural limit, documented as such.
        var (workbook, warnings) = Load(XlsxParts.Fixture(fixture));

        var tableWarnings = TableWarnings(warnings);
        await Assert.That(tableWarnings.Count).IsEqualTo(1);
        await Assert.That(tableWarnings[0].Subject).IsEqualTo(displayName);
        await Assert.That(tableWarnings[0].Detail).Contains("is not a valid table name");
        await Assert.That(workbook.Tables.Count).IsEqualTo(0);

        var formulaWarnings = warnings
            .Where(w => w.Kind == ExcelLoadWarningKind.UnparsableFormula)
            .ToList();
        await Assert.That(formulaWarnings.Count).IsEqualTo(1);
        await Assert.That(formulaWarnings[0].Subject).IsEqualTo("D1");
        await Assert.That(workbook.GetCellValue("Data", "D1").ToDouble()).IsEqualTo(42.0);
    }

    // === Malformed parts (schema-valid, so the loader is the only gate) ==================================

    [Test]
    [Arguments("name=\"T\" displayName=\"T\" ref=\"NOT-A-RANGE\"", "Item", "Valor", "T")]
    [Arguments("name=\"T\" displayName=\"T\"", "Item", "Valor", "T")]
    [Arguments(
        "name=\"T\" displayName=\"T\" ref=\"A1:B3\" headerRowCount=\"2\"",
        "Item",
        "Valor",
        "T"
    )]
    [Arguments(
        "name=\"T\" displayName=\"T\" ref=\"A1:B3\" totalsRowCount=\"2\"",
        "Item",
        "Valor",
        "T"
    )]
    [Arguments(
        "name=\"T\" displayName=\"T\" ref=\"A1:B1\" totalsRowCount=\"1\"",
        "Item",
        "Valor",
        "T"
    )]
    [Arguments("name=\"T\" displayName=\"T\" ref=\"A1:B3\"", "Item", "", "T")]
    [Arguments("name=\"T\" displayName=\"T\" ref=\"A1:B3\"", "c", "C", "T")]
    [Arguments("ref=\"A1:B3\"", "Item", "Valor", "Data")]
    public async Task Load_MalformedTablePart_WarnsOnce_SkipsTheTable_AndKeepsTheCells(
        string attributes,
        string firstColumn,
        string secondColumn,
        string expectedSubject
    )
    {
        // In order: a ref that is not an A1 range; no ref at all; two header rows (Excel writes 0 or 1);
        // two totals rows; header + totals over a single row; an empty column name; duplicate names under
        // OrdinalIgnoreCase; and a part with neither displayName nor name, whose Subject is the sheet.
        var path = WriteClosedXmlTable();

        try
        {
            OverwriteTablePart(
                path,
                TablePart(
                    attributes,
                    $"<tableColumn id=\"1\" name=\"{firstColumn}\"/><tableColumn id=\"2\" name=\"{secondColumn}\"/>"
                )
            );

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert
                .That(warnings[0].Kind)
                .IsEqualTo(ExcelLoadWarningKind.InvalidTableDefinition);
            await Assert.That(warnings[0].Subject).IsEqualTo(expectedSubject);
            await Assert.That(warnings[0].Detail).IsNotEmpty();
            await Assert.That(workbook.Tables.Count).IsEqualTo(0);
            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    [Arguments(
        "<tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Valor\"/><tableColumn id=\"3\" name=\"Extra\"/>"
    )]
    [Arguments("<tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\"/>")]
    [Arguments("")]
    public async Task Load_ColumnListDisagreeingWithTheRef_WarnsAndSkips(string columns)
    {
        // Three columns over a two-column ref; a <tableColumn> with no name attribute; no columns at all.
        var path = WriteClosedXmlTable();

        try
        {
            OverwriteTablePart(
                path,
                TablePart("name=\"T\" displayName=\"T\" ref=\"A1:B3\"", columns)
            );

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert
                .That(warnings[0].Kind)
                .IsEqualTo(ExcelLoadWarningKind.InvalidTableDefinition);
            await Assert.That(warnings[0].Subject).IsEqualTo("T");
            await Assert.That(workbook.Tables.Count).IsEqualTo(0);
            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    [Arguments("this is not xml")]
    [Arguments("<?xml version=\"1.0\"?><foo/>")]
    public async Task Load_UnreadableTablePart_WarnsWithTheSheetAsSubject(string xml)
    {
        // Garbage (XmlException) and a wrong root element (InvalidDataException) — the two failure modes
        // this reader introduces, since nothing touched the part before. The load survives; the Subject
        // is the sheet, because there is no displayName to be had.
        var path = WriteClosedXmlTable();

        try
        {
            OverwriteTablePart(path, xml);

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert
                .That(warnings[0].Kind)
                .IsEqualTo(ExcelLoadWarningKind.InvalidTableDefinition);
            await Assert.That(warnings[0].Subject).IsEqualTo("Data");
            await Assert.That(workbook.Tables.Count).IsEqualTo(0);
            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Shapes that ARE accepted ========================================================================

    [Test]
    public async Task Load_TablePartWithoutDisplayName_FallsBackToItsName()
    {
        // The schema requires displayName and makes name optional; a producer that wrote only name still
        // gets its table registered under it.
        var path = WriteClosedXmlTable();

        try
        {
            OverwriteTablePart(
                path,
                TablePart(
                    "name=\"OnlyName\" ref=\"A1:B3\"",
                    "<tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Valor\"/>"
                )
            );

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings).IsEmpty();
            await AssertTable(
                workbook.Tables["OnlyName"],
                "OnlyName",
                "Data",
                1,
                3,
                1,
                true,
                false,
                "Item",
                "Valor"
            );
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_DollarAnchoredRef_IsAccepted()
    {
        // Not a shape Excel writes, but the A1 parser behind DefineTable accepts the anchors, so the
        // loader does too rather than inventing a stricter rule.
        var path = WriteClosedXmlTable();

        try
        {
            OverwriteTablePart(
                path,
                TablePart(
                    "name=\"T\" displayName=\"T\" ref=\"$A$1:$B$3\"",
                    "<tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Valor\"/>"
                )
            );

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings).IsEmpty();
            await AssertTable(
                workbook.Tables["T"],
                "T",
                "Data",
                1,
                3,
                1,
                true,
                false,
                "Item",
                "Valor"
            );
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_ColumnNameEscapes_DecodeControlCharactersOnly()
    {
        // The SDK already decodes XML entities (a second pass would corrupt `&amp;` → `&` → …), and it
        // leaves OOXML's _xHHHH_ form alone. The reader decodes ONLY the code points a producer is FORCED
        // to escape (< 0x20, plus _x005f_ for the underscore that starts an escape), so a user-typed
        // literal `_x0020_`, a wrong-length hex run or a non-hex run stays exactly as typed.
        var path = WriteClosedXmlTable();

        try
        {
            OverwriteTablePart(
                path,
                TablePart(
                    "name=\"T\" displayName=\"T\" ref=\"A1:F3\"",
                    "<tableColumn id=\"1\" name=\"A&amp;B &lt; C &quot;q&quot; &#39;s\"/>"
                        + "<tableColumn id=\"2\" name=\"Tab_x0009_Here\"/>"
                        + "<tableColumn id=\"3\" name=\"Under_x005F_x\"/>"
                        + "<tableColumn id=\"4\" name=\"Keep_x0020_Space\"/>"
                        + "<tableColumn id=\"5\" name=\"Short_x00a_Run\"/>"
                        + "<tableColumn id=\"6\" name=\"Not_xZZZZ_Hex\"/>"
                )
            );

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings).IsEmpty();
            await AssertTable(
                workbook.Tables["T"],
                "T",
                "Data",
                1,
                3,
                1,
                true,
                false,
                "A&B < C \"q\" 's",
                "Tab\tHere",
                "Under_x",
                "Keep_x0020_Space",
                "Short_x00a_Run",
                "Not_xZZZZ_Hex"
            );
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Name collisions =================================================================================

    [Test]
    public async Task Load_DuplicateTableName_KeepsTheFirstRegistration_AndWarns()
    {
        // Excel forbids two tables sharing a name, so this is a corrupt file; the loader refuses to repoint
        // a name that formulas may already resolve through (first registration wins, in relationship
        // order: table1.xml = A1:B3 before table2.xml = F1:G3) and says so, rather than letting
        // DefineTable's redefine-replaces rule silently pick the LAST one.
        var path = WriteClosedXmlTable(extra: data =>
        {
            data.Cell("F1").Value = "K";
            data.Cell("G1").Value = "V";
            data.Cell("F2").Value = "k1";
            data.Cell("G2").Value = 5;
            data.Range("F1:G2").CreateTable("Tabela2");
        });

        try
        {
            OverwriteTablePart(
                path,
                TablePart(
                    "name=\"Tabela1\" displayName=\"Tabela1\" ref=\"F1:G2\"",
                    "<tableColumn id=\"1\" name=\"K\"/><tableColumn id=\"2\" name=\"V\"/>"
                ),
                containing: "Tabela2"
            );

            var (workbook, warnings) = Load(path);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert
                .That(warnings[0].Kind)
                .IsEqualTo(ExcelLoadWarningKind.InvalidTableDefinition);
            await Assert.That(warnings[0].Subject).IsEqualTo("Tabela1");
            await Assert.That(warnings[0].Detail).Contains("already");
            await Assert.That(workbook.Tables.Count).IsEqualTo(1);
            await AssertTable(
                workbook.Tables["Tabela1"],
                "Tabela1",
                "Data",
                1,
                3,
                1,
                true,
                false,
                "Item",
                "Valor"
            );
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_TableNameAlreadyUsedByADefinedName_TableWins_TheNameWarns()
    {
        // Tables and defined names share one namespace (Workbook.ThrowIfNameTaken). The sheets — and so
        // the tables — are read BEFORE the defined names, so on a corrupt file carrying both it is the
        // defined name that loses, with an InvalidDefinedName rather than an InvalidTableDefinition.
        var path = WriteClosedXmlTable(tableName: "Vendas");

        using (var document = SpreadsheetDocument.Open(path, isEditable: true))
        {
            var xlsxWorkbook = document.WorkbookPart!.Workbook!;
            var definedNames =
                xlsxWorkbook.DefinedNames
                ?? xlsxWorkbook.InsertAfter(new DefinedNames(), xlsxWorkbook.Sheets!);

            definedNames.AppendChild(new DefinedName("Data!$A$1:$A$3") { Name = "Vendas" });
            xlsxWorkbook.Save();
        }

        try
        {
            var (workbook, warnings) = Load(path);

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Kind).IsEqualTo(ExcelLoadWarningKind.InvalidDefinedName);
            await Assert.That(warnings[0].Subject).IsEqualTo("Vendas");
            await Assert.That(workbook.Tables.ContainsKey("Vendas")).IsTrue();
            await Assert.That(workbook.DefinedNames.ContainsKey("Vendas")).IsFalse();
        }
        finally
        {
            File.Delete(path);
        }
    }
}
