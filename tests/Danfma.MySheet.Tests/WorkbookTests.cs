using MemoryPack;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests;

public class WorkbookTests
{
    [Test]
    public async Task Workbook_IsSerializable()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Number(1);
        sheet["B1"] = Number(2);
        sheet["C1"] = Sum(Cell("A1", sheet), Cell("A2", sheet));

        var serialized = MemoryPackSerializer.Serialize(workbook);
        var deserialized = MemoryPackSerializer.Deserialize<Workbook>(serialized);

        await Assert.That(deserialized).IsEquivalentTo(workbook);
    }
}
