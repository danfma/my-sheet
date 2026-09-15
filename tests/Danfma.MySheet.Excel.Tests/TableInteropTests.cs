using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Interop with Excel <b>Tables</b> (a <c>&lt;table&gt;</c> part, a.k.a. a ListObject) and the STRUCTURED
/// REFERENCES they enable (<c>Tabela1[Valor]</c>). The loader reads each <c>&lt;table&gt;</c> part into
/// MySheet's table registry (<see cref="TableDefinitionReaderTests"/> pins the record it builds), so an
/// in-scope structured reference in a loaded file now EVALUATES against the loaded geometry, while the
/// out-of-scope forms (the current-row <c>[@Col]</c> / <c>[[#This Row],[Col]]</c> and an implicit-table
/// <c>[Col]</c>) still throw at the parse. What these tests pin is that a formula the load
/// cannot represent degrades the AFFECTED CELL ONLY (falling back to the cached value Excel stored
/// alongside it, reported via <see cref="ExcelLoadOptions.OnWarning"/>) instead of aborting the whole load.
///
/// ClosedXML writes the table for the hand-written fixtures below (an independent implementation); the
/// structured-reference formula cells are injected through the OpenXML SDK afterwards because ClosedXML
/// validates formulas on write and has no API for planting an arbitrary one. The committed <c>f*</c>
/// fixtures under Fixtures/ are the exception both ways: Aspose.Cells authored their tables AND typed
/// their formulas natively (that is what makes them the oracle fixtures).
/// </summary>
public class TableInteropTests
{
    // The in-scope spelling, kept on the ONE test that is about a structured reference itself.
    private const string StructuredFormula = "SUM(Tabela1[Valor])";

    // The vehicle for every test that only needs SOME formula the load rejects — the negative-parse cache,
    // the <v>-decode fallbacks and the two shared-master paths are not about tables at all. It is the
    // current-row item list a real file STORES for a typed `SUM(Tabela1[@Valor])`, and S1 keeps the
    // current-row forms permanently out of scope, so this vehicle cannot rot the way `Tabela1[Valor]` did.
    //
    // The spelling is measured on the surface that matters HERE, which is the sheet XML this loader reads, and
    // the two surfaces disagree — name the surface or the next reader will "correct" this note. Aspose.Cells
    // 26.6.0, PLAIN entry, 2026-09-11, `SUM(Tabela1[@Valor])` typed, saved, then unzipped: the raw
    // `<f>` reads `SUM(Tabela1[[#This Row],[Valor]])`, while Aspose's OWN object model reports
    // `=SUM(Tabela1[@Valor])` for the same cell after a reopen. Typing the item-list spelling instead
    // produces the identical `<f>`, so both spellings converge on the one this constant carries.
    private const string UnsupportedStructuredFormula = "SUM(Tabela1[[#This Row],[Valor]])";

    /// <summary>A "Data" sheet holding <c>Tabela1</c> over A1:B3 (header + two rows), plus whatever
    /// <paramref name="inject"/> plants into the raw sheet XML.</summary>
    private static string WriteTableFixture(Action<SheetData>? inject = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-table-{Guid.NewGuid():N}.xlsx");

        using (var fixture = new XLWorkbook())
        {
            var data = fixture.AddWorksheet("Data");

            data.Cell("A1").Value = "Item";
            data.Cell("B1").Value = "Valor";
            data.Cell("A2").Value = "a";
            data.Cell("B2").Value = 10;
            data.Cell("A3").Value = "b";
            data.Cell("B3").Value = 32;
            data.Range("A1:B3").CreateTable("Tabela1");

            fixture.SaveAs(path);
        }

        if (inject is not null)
        {
            using var document = SpreadsheetDocument.Open(path, isEditable: true);
            var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;

            inject(worksheet.GetFirstChild<SheetData>()!);
            worksheet.Save();
        }

        return path;
    }

    /// <summary>
    /// The A1:B3 fixture extended with a totals row: the part becomes <c>ref="A1:B4" totalsRowCount="1"</c>
    /// with the totals in row 4. ClosedXML writes the totals cell as <c>SUBTOTAL(109,[Valor])</c> with NO
    /// cached value (measured), so a 999 lie is cached over it here — keeping the cache-vs-evaluate
    /// distinction the assertions below need (the producer's own 42 would coincide with the true answer).
    /// </summary>
    private static string WriteTotalsTableFixture(Action<SheetData>? inject = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-totals-{Guid.NewGuid():N}.xlsx");

        using (var fixture = new XLWorkbook())
        {
            var data = fixture.AddWorksheet("Data");

            data.Cell("A1").Value = "Item";
            data.Cell("B1").Value = "Valor";
            data.Cell("A2").Value = "a";
            data.Cell("B2").Value = 10;
            data.Cell("A3").Value = "b";
            data.Cell("B3").Value = 32;

            var table = data.Range("A1:B3").CreateTable("Tabela1");

            table.ShowTotalsRow = true;
            table.Field("Valor").TotalsRowFunction = XLTotalsRowFunction.Sum;

            fixture.SaveAs(path);
        }

        using (var document = SpreadsheetDocument.Open(path, isEditable: true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var sheetData = worksheet.GetFirstChild<SheetData>()!;

            var row4 = sheetData.Elements<Row>().First(row => row.RowIndex?.Value == 4);
            var b4 = row4.Elements<Cell>().First(cell => cell.CellReference?.Value == "B4");

            b4.RemoveAllChildren();
            b4.AppendChild(new CellFormula("SUBTOTAL(109,[Valor])"));
            b4.AppendChild(new CellValue("999"));

            inject?.Invoke(sheetData);
            worksheet.Save();
        }

        return path;
    }

    // <c>&lt;f&gt;</c> before <c>&lt;v&gt;</c>, the order the schema requires. A null cachedValue omits
    // <c>&lt;v&gt;</c> entirely (a formula cell with no cached result at all); an empty string writes
    // <c>&lt;v/&gt;</c>. <paramref name="type"/> sets @t, so a cached result can be a string, an error or a
    // shared-string index rather than a number.
    private static Cell FormulaCell(
        string reference,
        CellFormula formula,
        string? cachedValue,
        CellValues? type = null
    )
    {
        var cell = new Cell { CellReference = reference };

        if (type is { } dataType)
        {
            cell.DataType = dataType;
        }

        cell.AppendChild(formula);

        if (cachedValue is not null)
        {
            cell.AppendChild(new CellValue(cachedValue));
        }

        return cell;
    }

    // Rows 1-3 come from ClosedXML; injected rows are appended in ascending order after them.
    private static void AppendToRow(SheetData sheetData, uint rowIndex, params Cell[] cells)
    {
        var row =
            sheetData
                .Elements<Row>()
                .FirstOrDefault(candidate => candidate.RowIndex?.Value == rowIndex)
            ?? sheetData.AppendChild(new Row { RowIndex = rowIndex });

        foreach (var cell in cells)
        {
            row.AppendChild(cell);
        }
    }

    [Test]
    public async Task Load_Table_RegistersItInWorkbookTables_WithoutWarnings()
    {
        // Baseline: the presence of a table part is harmless on its own — its cells are ordinary cells.
        // The cached <v> is a DELIBERATE LIE (999 != 10 + 32) so the assertion below can only pass if the
        // formula was really re-evaluated; a cache equal to the true sum would pin nothing.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula("SUM(B2:B3)"), "999"))
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(warnings.Count).IsEqualTo(0);
            await Assert.That(workbook.GetCellValue("Data", "B2").ToDouble()).IsEqualTo(10.0);
            // Re-evaluated by MySheet, not read from the cached <v>.
            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(42.0);
            // The loader registers the <table> part in the table registry (TableDefinitionReaderTests pins
            // the record it builds, geometry included), and it does NOT turn the table's name into a
            // defined name — the two registries stay distinct, exactly as Excel's Name Manager keeps them.
            await Assert.That(workbook.Tables.Count).IsEqualTo(1);
            await Assert.That(workbook.Tables.ContainsKey("Tabela1")).IsTrue();
            await Assert
                .That(string.Join("|", workbook.Tables["Tabela1"].ColumnNames))
                .IsEqualTo("Item|Valor");
            await Assert.That(workbook.DefinedNames.ContainsKey("Tabela1")).IsFalse();
        }
        finally
        {
            File.Delete(path);
        }
    }

    // The replacement for this class's old "reports a warning and falls back to 999" test, which asserted
    // exactly what the parser arm inverts, and for the interim "#NAME? with nothing registered" pin that
    // asserted what the loader reader inverted. It exists so the ARRIVAL of resolution is pinned rather
    // than hidden: nothing else in either suite would notice that `SUM(Tabela1[Valor])` started answering
    // the true sum.
    [Test]
    public async Task Load_StructuredReference_EvaluatesAgainstTheRegisteredTable()
    {
        // 999 is deliberately NOT the true sum of the Valor column (42), so the assertion below can only
        // pass on a real resolution: the pre-parser loader answered 999.0 (one UnparsableFormula warning),
        // the parser-without-reader interim answered #NAME? (no warning), and the oracle (Aspose.Cells
        // 26.6.0, PLAIN entry, same shape) answers 42.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula(StructuredFormula), "999"))
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            // The formula parsed AND the table registered: no warning at all, the cached 999 was not used.
            await Assert.That(warnings.Count).IsEqualTo(0);
            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(42.0);

            // The reference reached the dirty graph rather than being frozen at load time: editing a cell
            // under the column and recalculating moves the result.
            workbook["Data"]["B2"] = new Danfma.MySheet.Expressions.NumberValue(8);
            workbook.InvalidateCache();

            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(40.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_StructuredReference_WithOutOfScopeThisRowForms_StillDegradeWithAWarning()
    {
        // S1 keeps the current-row forms out of scope, so the multiplier shapes a real file carries
        // degrade their cells — BOTH spellings (the hand-typed `[@Valor]` and the item-list form the
        // producer actually stores for it, per Fixtures/README.md), and regardless of what surrounds the
        // reference. Asserting the ABSENCE of InvalidTableDefinition is the point: the table registered
        // fine, the formula SHAPE is what failed — the distinction the two warning kinds exist to make.
        var path = WriteTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                4,
                FormulaCell("B4", new CellFormula("Tabela1[@Valor]*2"), "999")
            );
            AppendToRow(
                sheetData,
                5,
                FormulaCell("B5", new CellFormula("Tabela1[[#This Row],[Valor]]*2"), "888")
            );
        });

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(warnings.Count).IsEqualTo(2);
            await Assert
                .That(
                    warnings.All(warning => warning.Kind == ExcelLoadWarningKind.UnparsableFormula)
                )
                .IsTrue();
            await Assert
                .That(warnings.Select(warning => warning.Subject).ToList())
                .IsEquivalentTo(new List<string> { "B4", "B5" });
            await Assert
                .That(
                    warnings.Any(warning =>
                        warning.Kind == ExcelLoadWarningKind.InvalidTableDefinition
                    )
                )
                .IsFalse();
            await Assert.That(workbook.Tables.Count).IsEqualTo(1);

            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(999.0);
            await Assert.That(workbook.GetCellValue("Data", "B5").ToDouble()).IsEqualTo(888.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_TableWithATotalsRow_ExcludesItFromTheDataBody()
    {
        // The part's ref INCLUDES the totals row (ref="A1:B4" totalsRowCount="1", while autoFilter stops
        // at A1:B3), so the one geometry mistake this shape invites — forgetting
        // lastDataRow = lastRow - totalsRows — produces a plausible number (10 + 32 + 999 = 1041) rather
        // than an error. Excel's rule being matched: Table[Column] is the data body only, [#All] is
        // header + data + totals, [#Totals] is the totals row itself.
        var path = WriteTotalsTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                5,
                FormulaCell("D1", new CellFormula("SUM(Tabela1[Valor])"), "888")
            );
            AppendToRow(
                sheetData,
                6,
                FormulaCell("D2", new CellFormula("ROWS(Tabela1[#All])"), "777")
            );
            AppendToRow(
                sheetData,
                7,
                FormulaCell("D3", new CellFormula("SUM(Tabela1[#Totals])"), "666")
            );
        });

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            // The totals cell's implicit [Valor] is S1 out-of-scope, so THAT cell degrades — to the 999
            // lie, not to a computed 42 — and nothing else warns.
            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Kind).IsEqualTo(ExcelLoadWarningKind.UnparsableFormula);
            await Assert.That(warnings[0].Subject).IsEqualTo("B4");

            // The data body stops short of the totals row: 42, not 1041.
            await Assert.That(workbook.GetCellValue("Data", "D1").ToDouble()).IsEqualTo(42.0);
            // [#All] spans header + data + totals.
            await Assert.That(workbook.GetCellValue("Data", "D2").ToDouble()).IsEqualTo(4.0);
            // [#Totals] reads the totals ROW — the cached 999, not a recomputed sum.
            await Assert.That(workbook.GetCellValue("Data", "D3").ToDouble()).IsEqualTo(999.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_StructuredReferenceFormula_WithNoCachedValue_LeavesTheCellBlank()
    {
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell("B4", new CellFormula(UnsupportedStructuredFormula), cachedValue: null)
            )
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Subject).IsEqualTo("B4");

            // Nothing to fall back to: the cell is blank rather than the load failing.
            await Assert
                .That(workbook.GetCellValue("Data", "B4").Kind)
                .IsEqualTo(ComputedValueKind.Blank);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_StructuredReferenceFormula_WithoutOptions_DoesNotThrow()
    {
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell("B4", new CellFormula(UnsupportedStructuredFormula), "999")
            )
        );

        try
        {
            // The plain overload: no callback, no observer — must still degrade instead of throwing.
            var workbook = ExcelFile.Load(path);

            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(999.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_RepeatedStructuredReferenceFormula_WarnsOncePerAffectedCell()
    {
        var path = WriteTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                4,
                FormulaCell("B4", new CellFormula(UnsupportedStructuredFormula), "999")
            );
            AppendToRow(
                sheetData,
                5,
                FormulaCell("B5", new CellFormula(UnsupportedStructuredFormula), "888")
            );
        });

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            // Identical text, but the SUBJECT is the cell — so each degraded cell is reported.
            await Assert.That(warnings.Count).IsEqualTo(2);
            await Assert
                .That(warnings.Select(warning => warning.Subject).ToList())
                .IsEquivalentTo(new List<string> { "B4", "B5" });

            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(999.0);
            await Assert.That(workbook.GetCellValue("Data", "B5").ToDouble()).IsEqualTo(888.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_SharedStructuredReferenceMaster_DoesNotAbort_AndTheGroupFallsBackToCachedValues()
    {
        // A dragged structured-reference formula: the master carries the text, the slave carries only si.
        // The master's text is rejected at the PARSE, not by the tokenizer (which reads the whole `[...]`
        // suffix as one token). What makes this a different code path from the plain-formula case above is
        // the shared GROUP: the rejected master must register nothing, so its slave degrades through the
        // missing-master fallback rather than through this cell's own catch.
        // The master's cached <v> is a DELIBERATE LIE: 999 is neither the column's true sum (10 + 32 = 42) nor
        // the oracle's #VALUE! for a current-row form outside the table, so the B4 assertion below can only
        // pass if the value came from the cache. The old cache, 42, coincided with the true sum.
        var path = WriteTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula(UnsupportedStructuredFormula)
                    {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0,
                        Reference = "B4:B5",
                    },
                    "999"
                )
            );
            AppendToRow(
                sheetData,
                5,
                FormulaCell(
                    "B5",
                    new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 },
                    "43"
                )
            );
        });

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            // One warning, for the master: the group's single formula text is what could not be parsed.
            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Kind).IsEqualTo(ExcelLoadWarningKind.UnparsableFormula);
            await Assert.That(warnings[0].Subject).IsEqualTo("B4");

            // The master falls back to its own cached value; the slave, whose master never registered,
            // falls back to its own through the existing missing-master path.
            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(999.0);
            await Assert.That(workbook.GetCellValue("Data", "B5").ToDouble()).IsEqualTo(43.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // --- The cached value the fallback reads is arbitrary producer output, not something Excel guarantees.
    // Every shape below used to abort the whole load: the degrade path decoded <v> with unguarded parses,
    // so "the load no longer dies on an unparsable formula" was only true when the cached value happened to
    // be a well-formed number. These pin the decode itself.

    [Test]
    public async Task Load_UnparsableFormula_WithEmptyCachedValue_LeavesTheCellBlank()
    {
        // <f>…</f><v/> — a formula cell the producer never evaluated. Not exotic: any tool that writes
        // formulas without computing them emits this.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell("B4", new CellFormula(UnsupportedStructuredFormula), "")
            )
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert
                .That(workbook.GetCellValue("Data", "B4").Kind)
                .IsEqualTo(ComputedValueKind.Blank);
            // The rest of the sheet still loaded — the point of the whole fallback.
            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
            await Assert
                .That(
                    warnings.Any(warning => warning.Kind == ExcelLoadWarningKind.UnparsableFormula)
                )
                .IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_UnparsableFormula_WithNonNumericCachedValue_FallsBackToTextAndWarns()
    {
        // No @t (so the numeric branch) but the text is not a number: a malformed cell, which must degrade
        // to text with a warning rather than throwing FormatException out of Load.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell("B4", new CellFormula(UnsupportedStructuredFormula), "abc")
            )
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(workbook.GetCellValue("Data", "B4").ToText()).IsEqualTo("abc");
            await Assert
                .That(
                    warnings.Any(warning =>
                        warning.Kind == ExcelLoadWarningKind.UnparsableCellLiteral
                        && warning.Subject == "B4"
                        && warning.Detail == "abc"
                    )
                )
                .IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_UnparsableFormula_WithCachedStringResult_FallsBackToThatText()
    {
        // t="str" is how Excel caches a formula whose RESULT is text — the normal shape for, say,
        // IF(Tabela1[Valor]>0,"ok","no").
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula(UnsupportedStructuredFormula),
                    "ok",
                    CellValues.String
                )
            )
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(workbook.GetCellValue("Data", "B4").ToText()).IsEqualTo("ok");
            // The fallback is not silent: the formula the parser rejected is reported against its cell.
            await Assert
                .That(
                    warnings.Any(warning =>
                        warning.Kind == ExcelLoadWarningKind.UnparsableFormula
                        && warning.Subject == "B4"
                    )
                )
                .IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_UnparsableFormula_WithCachedErrorResult_FallsBackToThatError()
    {
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula(UnsupportedStructuredFormula),
                    "#DIV/0!",
                    CellValues.Error
                )
            )
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });
            var value = workbook.GetCellValue("Data", "B4");

            await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
            await Assert.That(value.TryGetError(out var error)).IsTrue();
            await Assert.That(error.Display).IsEqualTo("#DIV/0!");
            // The fallback is not silent: the formula the parser rejected is reported against its cell.
            await Assert
                .That(
                    warnings.Any(warning =>
                        warning.Kind == ExcelLoadWarningKind.UnparsableFormula
                        && warning.Subject == "B4"
                    )
                )
                .IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_UnparsableFormula_WithOutOfRangeSharedStringIndex_LeavesTheCellBlankAndWarns()
    {
        // t="s" pointing past the end of the shared-string table: a stale index in a hand-edited file used
        // to escape as IndexOutOfRangeException.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula(UnsupportedStructuredFormula),
                    "9999",
                    CellValues.SharedString
                )
            )
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert
                .That(workbook.GetCellValue("Data", "B4").Kind)
                .IsEqualTo(ComputedValueKind.Blank);
            await Assert
                .That(
                    warnings.Any(warning =>
                        warning.Kind == ExcelLoadWarningKind.UnparsableCellLiteral
                        && warning.Subject == "B4"
                    )
                )
                .IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_SharedFormulaMasterRejectedAtParse_DegradesTheWholeGroup()
    {
        // INDEX takes 2-4 arguments, so this TOKENIZES fine and fails at the PARSE (InvalidArgumentCount) —
        // the only path where the group has already been seen by the tokenizer when the failure lands.
        const string BadArity = "INDEX(A1:C3,1,2,1,1)";

        var path = WriteTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula(BadArity)
                    {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0,
                        Reference = "B4:B5",
                    },
                    "111"
                )
            );
            AppendToRow(
                sheetData,
                5,
                FormulaCell(
                    "B5",
                    new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 },
                    "222"
                )
            );
        });

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Subject).IsEqualTo("B4");
            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(111.0);
            await Assert.That(workbook.GetCellValue("Data", "B5").ToDouble()).IsEqualTo(222.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_RejectedMasterReusingAnotherGroupsIndex_LeavesTheLegitimateGroupIntact()
    {
        // Out-of-spec (si is unique per worksheet in a valid file), but it is the shape that proves the
        // rejected master unregisters NOTHING: a naive "remove si on failure" would delete the healthy
        // group registered earlier under the same index, silently stripping ITS slaves' formulas.
        var path = WriteTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "C4",
                    new CellFormula("B2*2")
                    {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0,
                        Reference = "C4:C6",
                    },
                    "0"
                )
            );
            // A second "master" reusing si=0, whose text the parser rejects.
            AppendToRow(
                sheetData,
                5,
                FormulaCell(
                    "E5",
                    new CellFormula(UnsupportedStructuredFormula)
                    {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0,
                        Reference = "E5:E5",
                    },
                    "777"
                )
            );
            // A slave of the LEGITIMATE group, appearing after the rejected one.
            AppendToRow(
                sheetData,
                6,
                FormulaCell(
                    "C6",
                    new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 },
                    "0"
                )
            );
        });

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            // Only the rejected cell is reported.
            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Subject).IsEqualTo("E5");
            await Assert.That(workbook.GetCellValue("Data", "E5").ToDouble()).IsEqualTo(777.0);

            // The legitimate group kept its formulas: C4 = B2*2 = 20, and C6 is the slave two rows down,
            // so its shifted formula is B4*2. B4 is empty -> 0. What matters is that it is a FORMULA, not
            // the stale cached 0 — proven by editing B2 and B4 and recalculating.
            await Assert.That(workbook.GetCellValue("Data", "C4").ToDouble()).IsEqualTo(20.0);

            workbook["Data"]["B2"] = new Danfma.MySheet.Expressions.NumberValue(5);
            workbook["Data"]["B4"] = new Danfma.MySheet.Expressions.NumberValue(7);
            workbook.InvalidateCache();

            await Assert.That(workbook.GetCellValue("Data", "C4").ToDouble()).IsEqualTo(10.0);
            await Assert.That(workbook.GetCellValue("Data", "C6").ToDouble()).IsEqualTo(14.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_SharedMasterWithAnOverflowingRowNumber_DoesNotAbort()
    {
        // A row number past int.MaxValue makes the ANCHORED master parse throw OverflowException, not
        // ParseException — a different exception type reaching the same "this group is not anchored-safe"
        // decision. Excel cannot write such a row, but a hand-rolled producer can.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula("A99999999999+1")
                    {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0,
                        Reference = "B4:B5",
                    },
                    "1"
                )
            )
        );

        try
        {
            // Must not throw. The formula itself is nonsense and may evaluate to an error; the contract
            // pinned here is only that one bad cell cannot take the workbook down.
            var workbook = ExcelFile.Load(path);

            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // === Evaluation against the committed Aspose fixtures ================================================
    //
    // Every oracle number below is Aspose.Cells 26.6.0, PLAIN entry, as recorded in Fixtures/README.md —
    // the fixtures are the oracle's own output, so the assertions compare MySheet against it directly.

    private static (Workbook Workbook, List<ExcelLoadWarning> Warnings) LoadOracleFixture(
        string name
    )
    {
        var warnings = new List<ExcelLoadWarning>();
        var workbook = ExcelFile.Load(
            XlsxParts.Fixture(name),
            new ExcelLoadOptions { OnWarning = warnings.Add }
        );

        return (workbook, warnings);
    }

    [Test]
    public async Task Load_ItemSpecifiers_ResolveAgainstTheLoadedGeometry()
    {
        // Each specifier reads a DIFFERENT field of the geometry the loader derived from the part:
        // [#All] the whole ref, [#Data] the data band, [#Headers] the header row, and the composite
        // [[#Data],[Valor]] one column of that band — all four pinned from one real file.
        // Oracle (Fixtures/README.md): 3 / 2 / 2 / 42.
        var (workbook, warnings) = LoadOracleFixture("f1-plain");

        await Assert.That(warnings.Count).IsEqualTo(0);
        await Assert.That(workbook.GetCellValue("Data", "D2").ToDouble()).IsEqualTo(3.0);
        await Assert.That(workbook.GetCellValue("Data", "D3").ToDouble()).IsEqualTo(2.0);
        await Assert.That(workbook.GetCellValue("Data", "D4").ToDouble()).IsEqualTo(2.0);
        await Assert.That(workbook.GetCellValue("Data", "D5").ToDouble()).IsEqualTo(42.0);
    }

    [Test]
    public async Task Load_TableColumnNamesWithSpacesParenthesesAndAnApostrophe_ResolveFromFormulas()
    {
        // f4-names is the real-producer version of the escaping gauntlet: spaces and parentheses travel
        // verbatim, the apostrophe is stored RAW in the part but DOUBLED inside a formula's specifier
        // (Excel's escape — G4 was typed single and is still stored doubled), and the part's _x000a_
        // decodes to the newline the raw-newline specifier in G5 needs. A mismatch localizes to one
        // half: the raw registry names are pinned here, the results exercise the lexer and resolver.
        // Oracle (Fixtures/README.md): 42 / 2 / 3 / 3 / 300 / 15.
        var (workbook, warnings) = LoadOracleFixture("f4-names");

        await Assert.That(warnings.Count).IsEqualTo(0);
        await Assert
            .That(string.Join("|", workbook.Tables["Tabela1"].ColumnNames))
            .IsEqualTo(
                "AMOUNT IN USD For Line 1|(A) NAME OF PFIC|Owner's Share|Line\nBreak| Padded "
            );

        await Assert.That(workbook.GetCellValue("Data", "G1").ToDouble()).IsEqualTo(42.0);
        await Assert.That(workbook.GetCellValue("Data", "G2").ToDouble()).IsEqualTo(2.0);
        await Assert.That(workbook.GetCellValue("Data", "G3").ToDouble()).IsEqualTo(3.0);
        await Assert.That(workbook.GetCellValue("Data", "G4").ToDouble()).IsEqualTo(3.0);
        await Assert.That(workbook.GetCellValue("Data", "G5").ToDouble()).IsEqualTo(300.0);
        await Assert.That(workbook.GetCellValue("Data", "G6").ToDouble()).IsEqualTo(15.0);
    }

    [Test]
    public async Task Load_HeaderlessTable_ResolvesFromTheRealColumnNames()
    {
        // ref="A2:B3" headerRowCount="0" with row 1 empty — and the part keeps the REAL column names, so
        // the formulas resolve against the data band with no header row. SUM over the text column is 0,
        // and [#Headers] has nothing to span: the oracle still counts its #REF! as one COUNTA element
        // (measured, Fixtures/README.md) and MySheet matches. Oracle: 42 / 0 / 42 / 2 / 1 / 2.
        var (workbook, warnings) = LoadOracleFixture("f3-noheader");

        await Assert.That(warnings.Count).IsEqualTo(0);
        await Assert.That(workbook.GetCellValue("Data", "D1").ToDouble()).IsEqualTo(42.0);
        await Assert.That(workbook.GetCellValue("Data", "D2").ToDouble()).IsEqualTo(0.0);
        await Assert.That(workbook.GetCellValue("Data", "D3").ToDouble()).IsEqualTo(42.0);
        await Assert.That(workbook.GetCellValue("Data", "D4").ToDouble()).IsEqualTo(2.0);
        await Assert.That(workbook.GetCellValue("Data", "D5").ToDouble()).IsEqualTo(1.0);
        await Assert.That(workbook.GetCellValue("Data", "D6").ToDouble()).IsEqualTo(2.0);
    }

    [Test]
    public async Task Load_StructuredReference_FromAnotherSheet_ResolvesAgainstTheTable()
    {
        // The formulas sit on Report; the table lives on Data (TableDefinitionReaderTests pins the
        // SheetName side). The qualified spelling resolves identically — Aspose stores
        // `Data!Tabela1[Valor]` back WITHOUT the qualifier, so the unqualified form is the one the loader
        // actually meets. Oracle (Fixtures/README.md): 42 / 42.
        var (workbook, warnings) = LoadOracleFixture("f6-cross-sheet");

        await Assert.That(warnings.Count).IsEqualTo(0);
        await Assert.That(workbook.GetCellValue("Report", "D1").ToDouble()).IsEqualTo(42.0);
        await Assert.That(workbook.GetCellValue("Report", "D2").ToDouble()).IsEqualTo(42.0);
    }

    [Test]
    public async Task Load_DynamicStructuredReferenceThroughIndirect_Resolves()
    {
        // The runtime-constructed path: INDIRECT parses its own text with the expression parser, so the
        // table reference resolves only if the parser AND the resolver agree — and being volatile, the
        // recomputation must follow the text cell when it changes rather than stay frozen at the
        // load-time answer. D10 sits OUTSIDE the table (A1 is the table's own header row).
        // Oracle (Fixtures/README.md): 42; a text column sums to 0.
        var (workbook, warnings) = LoadOracleFixture("f1-plain");

        await Assert.That(warnings.Count).IsEqualTo(0);
        await Assert.That(workbook.GetCellValue("Data", "D6").ToDouble()).IsEqualTo(42.0);

        workbook["Data"]["D10"] = new Danfma.MySheet.Expressions.StringValue("Item");
        workbook.InvalidateCache();

        await Assert.That(workbook.GetCellValue("Data", "D6").ToDouble()).IsEqualTo(0.0);
    }

    [Test]
    public async Task Load_CurrentRowFormsFromARealFile_DegradeToTheirCachedValues()
    {
        // The shapes a real producer writes for typed `[@...]` input: D2 carries the item list inside
        // SUM, I2 as the whole cell. Both degrade to the cached values Aspose computed, while the
        // three-table sum on the same sheet evaluates — the degradation stays per-cell.
        // Oracle (Fixtures/README.md): D1 56; D2/I2 cached 10/20, current-row forms rejected.
        var (workbook, warnings) = LoadOracleFixture("f8-two-tables");

        await Assert
            .That(warnings.All(warning => warning.Kind == ExcelLoadWarningKind.UnparsableFormula))
            .IsTrue();
        await Assert
            .That(warnings.Select(warning => warning.Subject).ToList())
            .IsEquivalentTo(new List<string> { "D2", "I2" });

        await Assert.That(workbook.GetCellValue("Data", "D1").ToDouble()).IsEqualTo(56.0);
        await Assert.That(workbook.GetCellValue("Data", "D2").ToDouble()).IsEqualTo(10.0);
        await Assert.That(workbook.GetCellValue("Data", "I2").ToDouble()).IsEqualTo(20.0);
    }

    [Test]
    public async Task Load_HeaderOnlyTable_ItsStructuredReferencesAnswerTheEmptyReference_AcrossSaveAndLoad()
    {
        // Sweep item 33, end to end on the real file: f7 is Aspose-authored, ref="A1:B1" with a header row
        // and zero data rows, D1 = SUM(Tabela1[Valor]) and D2 = ROWS(Tabela1[#Data]) (oracle 0 / 0, both
        // entry modes). Before the empty-reference representation both answered #REF!. The cached <v> is
        // 0 as well, so the two rows alone could pass by reading the cache: the anchor rows below cannot —
        // Data!B2 = 7 sits directly under the header, OUTSIDE the table, so an explicit resize from the
        // anchor reads it (oracle, the loaded file with B2 = 7: 7) while the empty band never does (SUM 0).
        var (workbook, warnings) = LoadOracleFixture("f7-header-only");

        await Assert.That(warnings.Count).IsEqualTo(0);
        await Assert.That(workbook.GetCellValue("Data", "D1").AsObject()).IsEqualTo(0.0);
        await Assert.That(workbook.GetCellValue("Data", "D2").AsObject()).IsEqualTo(0.0);

        var data = workbook["Data"];
        data["B2"] = new Danfma.MySheet.Expressions.NumberValue(7);
        data["E1"] = Danfma.MySheet.Parsing.ExpressionParser.Parse(
            "=SUM(OFFSET(Tabela1[Valor],0,0,1,1))",
            data
        );
        data["E2"] = Danfma.MySheet.Parsing.ExpressionParser.Parse("=ISREF(Tabela1[Valor])", data);
        workbook.InvalidateCache();

        await Assert.That(workbook.GetCellValue("Data", "D1").AsObject()).IsEqualTo(0.0);
        await Assert.That(workbook.GetCellValue("Data", "E1").AsObject()).IsEqualTo(7.0);
        await Assert.That(workbook.GetCellValue("Data", "E2").AsObject() as bool?).IsTrue();

        // The representation is a runtime value: what the MemoryPack round trip carries is the header-only
        // table and the formulas, which resolve to the empty reference again after the load.
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path);
            var reloaded = Workbook.Load(path);

            await Assert.That(reloaded.Tables["Tabela1"].DataRowCount).IsEqualTo(0);
            await Assert.That(reloaded.GetCellValue("Data", "D1").AsObject()).IsEqualTo(0.0);
            await Assert.That(reloaded.GetCellValue("Data", "D2").AsObject()).IsEqualTo(0.0);
            await Assert.That(reloaded.GetCellValue("Data", "E1").AsObject()).IsEqualTo(7.0);
            await Assert.That(reloaded.GetCellValue("Data", "E2").AsObject() as bool?).IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }
}
