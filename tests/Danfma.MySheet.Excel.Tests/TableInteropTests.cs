using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// Interop with Excel <b>Tables</b> (a <c>&lt;table&gt;</c> part, a.k.a. a ListObject) and the STRUCTURED
/// REFERENCES they enable (<c>Tabela1[Valor]</c>). The loader does not populate MySheet's table registry —
/// <see cref="Workbook.Tables"/> exists and <c>Workbook.DefineTable</c> is its only writer, but nothing reads
/// an xlsx <c>&lt;table&gt;</c> part into it — and the tokenizer has no <c>[</c>, so a structured reference
/// cannot parse. What these tests pin is that such a formula degrades the AFFECTED CELL ONLY (falling back to
/// the cached value Excel stored alongside it, reported via <see cref="ExcelLoadOptions.OnWarning"/>) instead
/// of aborting the whole load.
///
/// ClosedXML writes the table (an independent implementation, like every other fixture here); the
/// structured-reference formula cells are injected through the OpenXML SDK afterwards because ClosedXML
/// validates formulas on write and has no API for planting an arbitrary one.
/// </summary>
public class TableInteropTests
{
    private const string StructuredFormula = "SUM(Tabela1[Valor])";

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
    public async Task Load_TableWithOrdinaryFormulas_LoadsAsAPlainRange_WithoutWarnings()
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
            // The loader does not populate the table registry, and it does not turn the table's name into a
            // defined name either — so nothing the evaluator resolves comes out of the <table> part.
            await Assert.That(workbook.Tables.Count).IsEqualTo(0);
            await Assert.That(workbook.DefinedNames.ContainsKey("Tabela1")).IsFalse();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_StructuredReferenceFormula_ReportsWarning_AndCellFallsBackToCachedValue()
    {
        // 999 is deliberately NOT the true sum of the Valor column (42), so the assertion can only pass
        // by reading the cached <v> — the behaviour under test — and not by coincidence.
        var path = WriteTableFixture(sheetData =>
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula(StructuredFormula), "999"))
        );

        try
        {
            var warnings = new List<ExcelLoadWarning>();

            var workbook = ExcelFile.Load(path, new ExcelLoadOptions { OnWarning = warnings.Add });

            await Assert.That(warnings.Count).IsEqualTo(1);
            await Assert.That(warnings[0].Kind).IsEqualTo(ExcelLoadWarningKind.UnparsableFormula);
            await Assert.That(warnings[0].Subject).IsEqualTo("B4");
            await Assert.That(warnings[0].Detail).IsNotEmpty();

            // The formula is gone, but the value Excel last computed for it survives.
            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(999.0);

            // Everything else loaded normally — this is a per-cell degradation, not a failed load.
            await Assert.That(workbook.GetCellValue("Data", "B3").ToDouble()).IsEqualTo(32.0);
            await Assert.That(workbook.GetCellValue("Data", "A2").ToText()).IsEqualTo("a");
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
                FormulaCell("B4", new CellFormula(StructuredFormula), cachedValue: null)
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
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula(StructuredFormula), "999"))
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
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula(StructuredFormula), "999"));
            AppendToRow(sheetData, 5, FormulaCell("B5", new CellFormula(StructuredFormula), "888"));
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
        // The master's TOKENIZATION is what fails here (before any parse), so this exercises a different
        // code path from the plain-formula case above.
        var path = WriteTableFixture(sheetData =>
        {
            AppendToRow(
                sheetData,
                4,
                FormulaCell(
                    "B4",
                    new CellFormula(StructuredFormula)
                    {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 0,
                        Reference = "B4:B5",
                    },
                    "42"
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
            await Assert.That(workbook.GetCellValue("Data", "B4").ToDouble()).IsEqualTo(42.0);
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
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula(StructuredFormula), ""))
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
            AppendToRow(sheetData, 4, FormulaCell("B4", new CellFormula(StructuredFormula), "abc"))
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
                FormulaCell("B4", new CellFormula(StructuredFormula), "ok", CellValues.String)
            )
        );

        try
        {
            var workbook = ExcelFile.Load(path);

            await Assert.That(workbook.GetCellValue("Data", "B4").ToText()).IsEqualTo("ok");
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
                FormulaCell("B4", new CellFormula(StructuredFormula), "#DIV/0!", CellValues.Error)
            )
        );

        try
        {
            var workbook = ExcelFile.Load(path);
            var value = workbook.GetCellValue("Data", "B4");

            await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Error);
            await Assert.That(value.TryGetError(out var error)).IsTrue();
            await Assert.That(error.Display).IsEqualTo("#DIV/0!");
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
                    new CellFormula(StructuredFormula),
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
        // INDEX takes 2-3 arguments, so this TOKENIZES fine and fails at the PARSE (InvalidArgumentCount) —
        // the only path where the group has already been seen by the tokenizer when the failure lands.
        const string BadArity = "INDEX(A1:C3,1,2,1)";

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
            // A second "master" reusing si=0, whose text cannot be tokenized at all.
            AppendToRow(
                sheetData,
                5,
                FormulaCell(
                    "E5",
                    new CellFormula(StructuredFormula)
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
}
