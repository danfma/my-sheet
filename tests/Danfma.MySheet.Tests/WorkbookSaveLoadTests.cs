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

            await Assert.That(loaded["Sheet1"]["A3"].Compute(loaded) as double?).IsEqualTo(3.0);
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

            await Assert.That(loaded["Sheet1"]["A3"].Compute(loaded) as double?).IsEqualTo(3.0);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
