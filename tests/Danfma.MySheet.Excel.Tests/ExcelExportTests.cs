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
