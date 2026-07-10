using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests;

public class WorkbookSaveLoadTests
{
    private static Workbook Sample()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);
        return workbook;
    }

    [Test]
    public async Task SaveAndLoad_RoundTrips()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            Sample().Save(path);
            var loaded = Workbook.Load(path);

            await Assert
                .That(loaded["Sheet1"]["A3"].Evaluate(loaded).AsObject() as double?)
                .IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task SaveAsyncAndLoadAsync_RoundTrips()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            await Sample().SaveAsync(path);
            var loaded = await Workbook.LoadAsync(path);

            await Assert
                .That(loaded["Sheet1"]["A3"].Evaluate(loaded).AsObject() as double?)
                .IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task SaveAndLoad_PreservesDefinedNames()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(10);
        sheet["A2"] = new NumberValue(20);
        sheet["A3"] = new NumberValue(30);
        workbook.DefineName("Vendas", "Sheet1!A1:A3");
        workbook.DefineName("Taxa", new NumberValue(0.1));

        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            workbook.Save(path);
            var loaded = Workbook.Load(path);

            // The names survive, keep their case-insensitive lookup, and re-evaluate to the same results.
            await Assert.That(loaded.DefinedNames.Count).IsEqualTo(2);
            await Assert
                .That(
                    ExpressionParser.Parse("=SUM(vendas)", sheet).Evaluate(loaded).AsObject()
                        as double?
                )
                .IsEqualTo(60.0);
            await Assert
                .That(
                    ExpressionParser.Parse("=Taxa*100", sheet).Evaluate(loaded).AsObject()
                        as double?
                )
                .IsEqualTo(10.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
