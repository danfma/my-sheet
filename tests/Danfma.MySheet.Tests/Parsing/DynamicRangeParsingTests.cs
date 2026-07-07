using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class DynamicRangeParsingTests
{
    private static Sheet NewSheet()
    {
        var workbook = new Workbook();
        return workbook.Sheets.Add("Sheet1");
    }

    [Test]
    public async Task IndexEndpoint_ParsesToDynamicRange()
    {
        var expr = ExpressionParser.Parse("=INDEX($D:$D,2):$D1", NewSheet());
        await Assert.That(expr).IsTypeOf<DynamicRange>();
    }

    [Test]
    public async Task StaticRange_StillParsesToRangeReference()
    {
        var expr = ExpressionParser.Parse("=A1:B2", NewSheet());
        await Assert.That(expr).IsTypeOf<RangeReference>();
    }

    [Test]
    public async Task WholeColumn_StillParsesToOpenRange()
    {
        var expr = ExpressionParser.Parse("=$D:$D", NewSheet());
        await Assert.That(expr).IsTypeOf<OpenRangeReference>();
    }
}
