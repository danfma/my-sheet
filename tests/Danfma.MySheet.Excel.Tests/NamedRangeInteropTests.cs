using ClosedXML.Excel;
using Danfma.MySheet;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Excel.Tests;

/// <summary>
/// xlsx interop for workbook-level defined names. ClosedXML (an independent implementation) is the oracle
/// on both sides: it writes the fixtures our reader loads, and reads back what our writer produced. Only
/// workbook-scoped, non-builtin names cross the boundary — sheet-scoped names (with a localSheetId) and the
/// builtin <c>_xlnm.*</c> names (e.g. Print_Area) are skipped by design.
/// </summary>
public class NamedRangeInteropTests
{
    private static async Task WithFixture(Action<XLWorkbook> build, Func<Workbook, Task> assert)
    {
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-names-{Guid.NewGuid():N}.xlsx");

        try
        {
            using (var fixture = new XLWorkbook())
            {
                build(fixture);
                fixture.SaveAs(path);
            }

            await assert(ExcelFile.Load(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task Load_WorkbookScopedName_ResolvesInFormula()
    {
        await WithFixture(
            fixture =>
            {
                var data = fixture.AddWorksheet("Data");
                data.Cell("A1").Value = 10;
                data.Cell("A2").Value = 20;
                data.Cell("A3").Value = 30;
                data.Range("A1:A3").AddToNamed("Vendas", XLScope.Workbook);
            },
            async workbook =>
            {
                await Assert.That(workbook.DefinedNames.ContainsKey("Vendas")).IsTrue();

                // The imported name expands in a formula, just like a native defined name would.
                var result = ExpressionParser
                    .Parse("=SUM(Vendas)", workbook["Data"])
                    .Evaluate(workbook)
                    .AsObject();

                await Assert.That(result as double?).IsEqualTo(60.0);
            }
        );
    }

    [Test]
    public async Task Load_SkipsSheetScopedAndBuiltinNames()
    {
        await WithFixture(
            fixture =>
            {
                var data = fixture.AddWorksheet("Data");
                data.Cell("A1").Value = 1;
                data.Cell("A2").Value = 2;

                data.Range("A1:A2").AddToNamed("GlobalName", XLScope.Workbook);
                data.Range("A1:A2").AddToNamed("LocalName", XLScope.Worksheet); // sheet-scoped (localSheetId)
                data.PageSetup.PrintAreas.Add("A1:A2"); // builtin _xlnm.Print_Area
            },
            async workbook =>
            {
                // Only the workbook-scoped, non-builtin name is imported; the rest are skipped without error.
                await Assert.That(workbook.DefinedNames.ContainsKey("GlobalName")).IsTrue();
                await Assert.That(workbook.DefinedNames.ContainsKey("LocalName")).IsFalse();
                await Assert
                    .That(
                        workbook.DefinedNames.Keys.Any(key =>
                            key.Contains("Print", StringComparison.OrdinalIgnoreCase)
                        )
                    )
                    .IsFalse();
                await Assert.That(workbook.DefinedNames.Count).IsEqualTo(1);
            }
        );
    }

    [Test]
    public async Task Save_WritesFullyQualifiedName_ReadableByClosedXml()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new NumberValue(10);
        data["A2"] = new NumberValue(20);
        data["A3"] = new NumberValue(30);
        workbook.DefineName("Vendas", "Data!A1:A3");

        var path = Path.Combine(Path.GetTempPath(), $"mysheet-names-{Guid.NewGuid():N}.xlsx");

        try
        {
            workbook.SaveAsExcel(path);

            using var oracle = new XLWorkbook(path);

            await Assert.That(oracle.DefinedNames.TryGetValue("Vendas", out var defined)).IsTrue();

            // The refersTo is fully qualified (empty un-parse context -> every reference keeps its sheet).
            await Assert.That(defined!.RefersTo).Contains("Data");
            await Assert.That(defined.RefersTo).Contains("A1");
            await Assert.That(defined.RefersTo).Contains("A3");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task RoundTrip_SaveThenLoad_PreservesName()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        data["A1"] = new NumberValue(10);
        data["A2"] = new NumberValue(20);
        data["A3"] = new NumberValue(30);
        workbook.DefineName("Vendas", "Data!A1:A3");

        var path = Path.Combine(Path.GetTempPath(), $"mysheet-names-{Guid.NewGuid():N}.xlsx");

        try
        {
            workbook.SaveAsExcel(path);
            var reloaded = ExcelFile.Load(path);

            await Assert.That(reloaded.DefinedNames.ContainsKey("Vendas")).IsTrue();

            var result = ExpressionParser
                .Parse("=SUM(Vendas)", reloaded["Data"])
                .Evaluate(reloaded)
                .AsObject();

            await Assert.That(result as double?).IsEqualTo(60.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
