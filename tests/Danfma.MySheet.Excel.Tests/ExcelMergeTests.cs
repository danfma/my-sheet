using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Excel.Tests;

public class ExcelMergeTests
{
    /// <summary>Target file: Data!A1=1 (literal), Data!A2 = formula A1*10, Data!B5 = "keep" (untouched
    /// cell), plus a whole sheet "Other" our workbook does not have.</summary>
    private static string CreateTargetFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-merge-{Guid.NewGuid():N}.xlsx");

        using var target = new XLWorkbook();
        var data = target.AddWorksheet("Data");
        data.Cell("A1").Value = 1;
        data.Cell("A2").FormulaA1 = "A1*10";
        data.Cell("B5").Value = "keep";
        target.AddWorksheet("Other").Cell("A1").Value = 99;
        target.SaveAs(path);

        return path;
    }

    /// <summary>Our values: Data!A2 = 5 (computed from a formula), Data!C1 text, Data!D1 an error, plus a
    /// sheet "Missing" that does not exist in the target (must be skipped).</summary>
    private static Workbook BuildWorkbook()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");

        data["A2"] = ExpressionParser.Parse("=2+3", data);
        data["C1"] = ExpressionParser.Parse("=\"te\"&\"xt\"", data);
        data["D1"] = ExpressionParser.Parse("=1/0", data);

        workbook.Sheets.Add("Missing")["A1"] = ExpressionParser.Parse(
            "=1",
            workbook.Sheets["Missing"]
        );

        return workbook;
    }

    [Test]
    public async Task Merge_InjectsLiterals_DropsTargetFormulas_PreservesEverythingElse()
    {
        var path = CreateTargetFile();

        try
        {
            BuildWorkbook().MergeIntoExcel(path);

            using var merged = new XLWorkbook(path);
            var data = merged.Worksheet("Data");

            // Our value landed as a literal and the target cell's formula was dropped.
            await Assert.That(data.Cell("A2").GetDouble()).IsEqualTo(5.0);
            await Assert.That(data.Cell("A2").HasFormula).IsFalse();

            // New cells were created where the target had none.
            await Assert.That(data.Cell("C1").GetString()).IsEqualTo("text");
            await Assert.That(data.Cell("D1").Value.IsError).IsTrue();
            await Assert.That(data.Cell("D1").Value.GetError()).IsEqualTo(XLError.DivisionByZero);

            // Everything we do not own is intact: other cells, other sheets; our extra sheet was skipped.
            await Assert.That(data.Cell("A1").GetDouble()).IsEqualTo(1.0);
            await Assert.That(data.Cell("B5").GetString()).IsEqualTo("keep");
            await Assert.That(merged.Worksheet("Other").Cell("A1").GetDouble()).IsEqualTo(99.0);
            await Assert.That(merged.Worksheets.Contains("Missing")).IsFalse();
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Merge_TemplateWorkflow_IsCopyThenMerge()
    {
        // The template→report flow is deliberately NOT an overload: copy the pristine template
        // yourself, then merge into the copy. This test documents the recipe.
        var templatePath = CreateTargetFile();
        var reportPath = Path.Combine(
            Path.GetTempPath(),
            $"mysheet-report-{Guid.NewGuid():N}.xlsx"
        );

        try
        {
            File.Copy(templatePath, reportPath);
            BuildWorkbook().MergeIntoExcel(reportPath);

            using var report = new XLWorkbook(reportPath);
            await Assert.That(report.Worksheet("Data").Cell("A2").GetDouble()).IsEqualTo(5.0);

            // The template itself still has its formula, untouched.
            using var original = new XLWorkbook(templatePath);
            await Assert.That(original.Worksheet("Data").Cell("A2").FormulaA1).IsEqualTo("A1*10");
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(reportPath);
        }
    }

    /// <summary>Target: Data!A1=1, Data!A2 formula, Skip!A1="orig".</summary>
    private static string CreateTargetWithSkipSheet()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-skip-{Guid.NewGuid():N}.xlsx");

        using var target = new XLWorkbook();
        var data = target.AddWorksheet("Data");
        data.Cell("A1").Value = 1;
        data.Cell("A2").FormulaA1 = "A1*10";
        target.AddWorksheet("Skip").Cell("A1").Value = "orig";
        target.SaveAs(path);

        return path;
    }

    /// <summary>Our values: Data!A2 = 5, Skip!A1 = 42 (the value that must NOT land when Skip is ignored).</summary>
    private static Workbook BuildWorkbookWithSkip()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A2"] = ExpressionParser.Parse("=2+3", data);

        var skip = workbook.Sheets.Add("Skip");
        skip["A1"] = ExpressionParser.Parse("=42", skip);

        return workbook;
    }

    [Test]
    public async Task Merge_WithIgnoredSheet_SkipsIt()
    {
        var path = CreateTargetWithSkipSheet();

        try
        {
            BuildWorkbookWithSkip().MergeIntoExcel(path, new HashSet<string> { "Skip" });

            using var merged = new XLWorkbook(path);

            // Non-ignored sheet merged as usual.
            await Assert.That(merged.Worksheet("Data").Cell("A2").GetDouble()).IsEqualTo(5.0);
            // Ignored sheet untouched: our 42 was NOT written, the target's "orig" stays.
            await Assert.That(merged.Worksheet("Skip").Cell("A1").GetString()).IsEqualTo("orig");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Merge_IgnoredSheet_IsCaseInsensitive()
    {
        var path = CreateTargetWithSkipSheet();

        try
        {
            // Lowercase "skip" must still skip the "Skip" sheet.
            BuildWorkbookWithSkip().MergeIntoExcel(path, new HashSet<string> { "skip" });

            using var merged = new XLWorkbook(path);

            await Assert.That(merged.Worksheet("Skip").Cell("A1").GetString()).IsEqualTo("orig");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // A target authored with formulas carries a calcChain (xl/calcChain.xml). Merging overrides/drops
    // formula cells, leaving that calcChain stale — Excel then reports "Removed records: Formula from
    // calcChain.xml" and shows a repair dialog. The merge must remove the stale calcChain so Excel
    // rebuilds it silently on open.
    [Test]
    public async Task Merge_PreservesStylesMergedRangesAndUntouchedContent()
    {
        // The streaming merge must copy everything it does not own verbatim: a cell's style index (so a
        // formatted template cell stays formatted), merged ranges, and untouched cells/rows.
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-preserve-{Guid.NewGuid():N}.xlsx");

        using (var target = new XLWorkbook())
        {
            var data = target.AddWorksheet("Data");
            data.Cell("A1").Value = 0; // we will overwrite this value…
            data.Cell("A1").Style.NumberFormat.Format = "0.00"; // …but this format must survive
            data.Cell("A1").Style.Font.Bold = true;
            data.Range("B2:D2").Merge(); // a merged range far from our cell
            data.Cell("F6").Value = "untouched";
            target.SaveAs(path);
        }

        try
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.Add("Data");
            sheet["A1"] = ExpressionParser.Parse("=2+2", sheet);

            workbook.MergeIntoExcel(path);

            using var merged = new XLWorkbook(path);
            var data = merged.Worksheet("Data");

            // Our value landed as a literal…
            await Assert.That(data.Cell("A1").GetDouble()).IsEqualTo(4.0);
            await Assert.That(data.Cell("A1").HasFormula).IsFalse();

            // …while the target cell's formatting (style index) was preserved.
            await Assert.That(data.Cell("A1").Style.NumberFormat.Format).IsEqualTo("0.00");
            await Assert.That(data.Cell("A1").Style.Font.Bold).IsTrue();

            // Merged ranges and untouched cells are intact.
            await Assert
                .That(data.MergedRanges.Select(r => r.RangeAddress.ToString()))
                .Contains("B2:D2");
            await Assert.That(data.Cell("F6").GetString()).IsEqualTo("untouched");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Merge_RemovesStaleCalcChain()
    {
        var path = CreateTargetFile(); // Data!A2 is a formula → the authored file has a calcChain

        try
        {
            using (var pre = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, false))
            {
                // Precondition: the authored file really carries a calcChain (else the test is vacuous).
                await Assert.That(pre.WorkbookPart!.CalculationChainPart).IsNotNull();
            }

            BuildWorkbook().MergeIntoExcel(path); // overrides Data!A2, dropping its formula

            using var merged = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(
                path,
                false
            );
            await Assert.That(merged.WorkbookPart!.CalculationChainPart).IsNull();
        }
        finally
        {
            File.Delete(path);
        }
    }
}
