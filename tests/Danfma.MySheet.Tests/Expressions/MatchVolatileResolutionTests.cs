using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class MatchVolatileResolutionTests
{
    [Test]
    public async Task Match_ResolvesAComputedLookupArrayOnce()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["B1"] = new NumberValue(10);
        sheet["B2"] = new NumberValue(20);
        sheet["B3"] = new NumberValue(30);
        var draws = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++draws);
        sheet["AZ5000"] = ExpressionParser.Parse("=MATCH(20,OFFSET(B1,TICK()-1,0,3,1),0)", sheet);

        var result = workbook.GetCellValue("Main", "AZ5000");
        var value = result.TryGetError(out var error)
            ? error.ToString()
            : result.ToDouble().ToString(CultureInfo.InvariantCulture);

        // Before this fix the three resolutions drew three times and returned #N/A instead of position 2.
        await Assert.That(value).IsEqualTo("2");
        await Assert.That(draws).IsEqualTo(1);
    }

    [Test]
    [Arguments("=XMATCH(1,IF(TICK()>0,{1;2}),0)", 1)]
    [Arguments("=XMATCH(TICK(),A1:A3,TICK())", 2)]
    public async Task XMatch_EvaluatesEachVolatileSlotOnce(string formula, int expectedDraws)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        var draws = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++draws);

        _ = ExpressionParser.Parse(formula, sheet).Evaluate(workbook);

        await Assert.That(draws).IsEqualTo(expectedDraws);
    }
}
