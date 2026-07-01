using ClosedXML.Excel;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Excel.Tests;

public class ExcelMergeTests
{
    /// <summary>Template: Data!A1=1 (literal), Data!A2 = formula A1*10, Data!B5 = "keep" (untouched cell),
    /// plus a whole sheet "Other" our workbook does not have.</summary>
    private static string CreateTemplate()
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-template-{Guid.NewGuid():N}.xlsx");

        using var template = new XLWorkbook();
        var data = template.AddWorksheet("Data");
        data.Cell("A1").Value = 1;
        data.Cell("A2").FormulaA1 = "A1*10";
        data.Cell("B5").Value = "keep";
        template.AddWorksheet("Other").Cell("A1").Value = 99;
        template.SaveAs(path);

        return path;
    }

    /// <summary>Our values: Data!A2 = 5 (computed from a formula), Data!C1 text, Data!D1 an error, plus a
    /// sheet "Missing" that does not exist in the template (must be skipped).</summary>
    private static Workbook BuildWorkbook()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");

        data["A2"] = ExpressionParser.Parse("=2+3", data);
        data["C1"] = ExpressionParser.Parse("=\"te\"&\"xt\"", data);
        data["D1"] = ExpressionParser.Parse("=1/0", data);

        workbook.Sheets.Add("Missing")["A1"] = ExpressionParser.Parse("=1", workbook.Sheets["Missing"]);

        return workbook;
    }

    [Test]
    public async Task Merge_NonDestructive_InjectsLiterals_PreservesEverythingElse()
    {
        var templatePath = CreateTemplate();
        var outputPath = Path.Combine(Path.GetTempPath(), $"mysheet-merged-{Guid.NewGuid():N}.xlsx");

        try
        {
            BuildWorkbook().MergeIntoExcel(templatePath, outputPath);

            using var merged = new XLWorkbook(outputPath);
            var data = merged.Worksheet("Data");

            // Our value landed as a literal and the target cell's formula was dropped.
            await Assert.That(data.Cell("A2").GetDouble()).IsEqualTo(5.0);
            await Assert.That(data.Cell("A2").HasFormula).IsFalse();

            // New cells were created where the template had none.
            await Assert.That(data.Cell("C1").GetString()).IsEqualTo("text");
            await Assert.That(data.Cell("D1").Value.IsError).IsTrue();
            await Assert.That(data.Cell("D1").Value.GetError()).IsEqualTo(XLError.DivisionByZero);

            // Everything we do not own is intact: other cells, other sheets, our missing sheet skipped.
            await Assert.That(data.Cell("A1").GetDouble()).IsEqualTo(1.0);
            await Assert.That(data.Cell("B5").GetString()).IsEqualTo("keep");
            await Assert.That(merged.Worksheet("Other").Cell("A1").GetDouble()).IsEqualTo(99.0);
            await Assert.That(merged.Worksheets.Contains("Missing")).IsFalse();

            // Non-destructive: the template itself still has its formula.
            using var original = new XLWorkbook(templatePath);
            await Assert.That(original.Worksheet("Data").Cell("A2").FormulaA1).IsEqualTo("A1*10");
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(outputPath);
        }
    }

    [Test]
    public async Task Merge_InPlace_EditsTheFileDirectly()
    {
        var path = CreateTemplate();

        try
        {
            BuildWorkbook().MergeIntoExcel(path);

            using var merged = new XLWorkbook(path);

            await Assert.That(merged.Worksheet("Data").Cell("A2").GetDouble()).IsEqualTo(5.0);
            await Assert.That(merged.Worksheet("Data").Cell("A2").HasFormula).IsFalse();
            await Assert.That(merged.Worksheet("Data").Cell("B5").GetString()).IsEqualTo("keep");
        }
        finally
        {
            File.Delete(path);
        }
    }
}
