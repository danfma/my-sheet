using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Excel.Tests;

public class ExcelExportTests
{
    /// <summary>Two sheets: literals in Data!A1/A2/B1/C1, formulas in Data!A3 (A1+A2), Data!D1 (1/0)
    /// and a cross-sheet Summary!A1 (Data!A3*2).</summary>
    private static Workbook BuildWorkbook()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");

        data["A1"] = new NumberValue(2);
        data["A2"] = new NumberValue(3);
        data["B1"] = new StringValue("hello");
        data["C1"] = new BooleanValue(true);
        data["A3"] = ExpressionParser.Parse("=A1+A2", data);
        data["D1"] = ExpressionParser.Parse("=1/0", data);

        var summary = workbook.Sheets.Add("Summary");
        summary["A1"] = ExpressionParser.Parse("=Data!A3*2", summary);

        return workbook;
    }

    private static async Task WithExport(
        ExcelExportOptions? options,
        Func<string, Task> assert
    )
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-export-{Guid.NewGuid():N}.xlsx");

        try
        {
            BuildWorkbook().SaveAsExcel(path, options);
            await assert(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task ValuesOnly_FlattensFormulasToLiterals()
    {
        await WithExport(
            new ExcelExportOptions { FormulaMode = FormulaMode.ValuesOnly },
            async path =>
            {
                var reloaded = ExcelFile.Load(path);

                // The formula cell came back as a literal (the snapshot), not as a formula.
                await Assert.That(reloaded["Data"]["A3"]).IsTypeOf<NumberValue>();
                await Assert.That(reloaded.GetCellValue("Data", "A3").ToDouble()).IsEqualTo(5.0);

                // Literals and sheet order survive.
                await Assert.That(reloaded["Data"].Index).IsEqualTo(0);
                await Assert.That(reloaded["Summary"].Index).IsEqualTo(1);
                await Assert.That(reloaded.GetCellValue("Data", "B1").ToText()).IsEqualTo("hello");
                await Assert.That(reloaded.GetCellValue("Data", "C1").TryGetBoolean(out var flag)).IsTrue();
                await Assert.That(flag).IsTrue();
                await Assert.That(reloaded.GetCellValue("Summary", "A1").ToDouble()).IsEqualTo(10.0);

                // The division error is written as an error literal.
                await Assert.That(reloaded.GetCellValue("Data", "D1").TryGetError(out var error)).IsTrue();
                await Assert.That(error).IsEqualTo(Error.DivZero);
            }
        );
    }

    [Test]
    public async Task ValuesOnly_FormulaWithEmptyResult_WritesZero()
    {
        // Excel-parity: a formula whose result is blank displays 0, so ValuesOnly writes 0 (it used to omit
        // the cell). =F10 with F10 empty round-trips as a literal 0, not a missing cell.
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-export-{Guid.NewGuid():N}.xlsx");

        try
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.Add("Data");
            sheet["A1"] = ExpressionParser.Parse("=F10", sheet);

            workbook.SaveAsExcel(path, new ExcelExportOptions { FormulaMode = FormulaMode.ValuesOnly });

            var reloaded = ExcelFile.Load(path);

            await Assert.That(reloaded["Data"]["A1"]).IsTypeOf<NumberValue>();
            await Assert.That(reloaded.GetCellValue("Data", "A1").ToDouble()).IsEqualTo(0.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Formulas_RoundTripThroughOurReader()
    {
        await WithExport(
            new ExcelExportOptions { FormulaMode = FormulaMode.Formulas },
            async path =>
            {
                var reloaded = ExcelFile.Load(path);

                // Formula cells come back as real expression trees and re-evaluate to the same results.
                await Assert.That(reloaded["Data"]["A3"]).IsTypeOf<BinaryOperation>();
                await Assert.That(reloaded.GetCellValue("Data", "A3").ToDouble()).IsEqualTo(5.0);
                await Assert.That(reloaded["Summary"]["A1"]).IsTypeOf<BinaryOperation>();
                await Assert.That(reloaded.GetCellValue("Summary", "A1").ToDouble()).IsEqualTo(10.0);
                await Assert.That(reloaded.GetCellValue("Data", "D1").TryGetError(out var error)).IsTrue();
                await Assert.That(error).IsEqualTo(Error.DivZero);

                // Plain literals stay literals.
                await Assert.That(reloaded["Data"]["A1"]).IsTypeOf<NumberValue>();
            }
        );
    }

    [Test]
    public async Task Formulas_WritesFormulaTextAndCachedValues_ClosedXmlOracle()
    {
        await WithExport(
            new ExcelExportOptions { FormulaMode = FormulaMode.Formulas },
            async path =>
            {
                using var oracle = new XLWorkbook(path);
                var data = oracle.Worksheet("Data");

                // The formula text is there (so Excel keeps recalculating the file)…
                await Assert.That(data.Cell("A3").FormulaA1).IsEqualTo("A1+A2");
                await Assert.That(oracle.Worksheet("Summary").Cell("A1").FormulaA1).IsEqualTo("Data!A3*2");

                // …and the cached value is written alongside, so viewers show it without recalc.
                await Assert.That(data.Cell("A3").CachedValue.GetNumber()).IsEqualTo(5.0);
            }
        );
    }

    [Test]
    public async Task Text_WithEdgeWhitespace_EmitsXmlSpacePreserve()
    {
        // Regression for the whitespace-trimming bug. Spec-compliant readers (Excel, Aspose.Cells) strip
        // leading/trailing whitespace off a shared-string <t> unless it carries xml:space="preserve".
        // ClosedXML is lenient and keeps the spaces regardless, so a reader-based round-trip is a false
        // negative — we must assert on the raw XML the file actually contains.
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-export-{Guid.NewGuid():N}.xlsx");

        try
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.Add("Data");
            sheet["A1"] = new StringValue("  leading");
            sheet["A2"] = new StringValue("trailing  ");
            sheet["A3"] = new StringValue("clean");
            sheet["A4"] = new StringValue("   ");

            workbook.SaveAsExcel(path, new ExcelExportOptions { FormulaMode = FormulaMode.ValuesOnly });

            var sharedStringsXml = ReadEntry(path, "xl/sharedStrings.xml");

            // Edge-whitespace strings are tagged so real spreadsheets keep the spaces...
            await Assert.That(sharedStringsXml).Contains("xml:space=\"preserve\">  leading</");
            await Assert.That(sharedStringsXml).Contains("xml:space=\"preserve\">trailing  </");
            await Assert.That(sharedStringsXml).Contains("xml:space=\"preserve\">   </");

            // ...while a clean string is left untagged (no attribute bloat).
            await Assert.That(sharedStringsXml).Contains("<x:t>clean</x:t>");

            // ClosedXML (lenient) and our own reader both keep the spaces, closing the loop.
            using var oracle = new XLWorkbook(path);
            var data = oracle.Worksheet("Data");
            await Assert.That(data.Cell("A1").GetString()).IsEqualTo("  leading");
            await Assert.That(data.Cell("A2").GetString()).IsEqualTo("trailing  ");
            await Assert.That(data.Cell("A4").GetString()).IsEqualTo("   ");

            var reloaded = ExcelFile.Load(path);
            await Assert.That(reloaded.GetCellValue("Data", "A1").ToText()).IsEqualTo("  leading");
            await Assert.That(reloaded.GetCellValue("Data", "A2").ToText()).IsEqualTo("trailing  ");
            await Assert.That(reloaded.GetCellValue("Data", "A4").ToText()).IsEqualTo("   ");
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string ReadEntry(string xlsxPath, string entryName)
    {
        using var archive = System.IO.Compression.ZipFile.OpenRead(xlsxPath);
        var entry =
            archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Entry '{entryName}' not found in {xlsxPath}.");
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    [Test]
    public async Task DefaultOptions_AreValuesOnly()
    {
        await WithExport(
            options: null,
            async path =>
            {
                var reloaded = ExcelFile.Load(path);

                await Assert.That(reloaded["Data"]["A3"]).IsTypeOf<NumberValue>();
            }
        );
    }
}
