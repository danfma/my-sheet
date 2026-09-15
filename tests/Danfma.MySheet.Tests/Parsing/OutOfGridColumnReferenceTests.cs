using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class OutOfGridColumnReferenceTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task OutOfGridColumnRangesAreRefErrors_AndBareTokensRemainNames()
    {
        // Empty fixture, formula at AZ5000 in Aspose.Cells 26.7.0: PLAIN/CSE both return 1 for each range,
        // including XFE1:XFE1, while the bare token returns #NAME?. MySheet deliberately rejects out-of-grid
        // ranges as #REF! but keeps an invalid bare cell-shaped token as a name.
        await Assert.That(Calc("=COUNTA(FXSHRXW1:FXSHRXW1)")).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=AAAAAAAAAAAAAAAAAAAAAAAA1")).IsEqualTo(ErrorValue.Name);
        await Assert
            .That(Calc("=COUNTA(AAAAAAAAAAAAAAAAAAAAAAAA1:AAAAAAAAAAAAAAAAAAAAAAAA5)"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=COUNTA(XFD1:XFD1)") as double?).IsEqualTo(0d);
        await Assert.That(Calc("=COUNTA(XFE1:XFE1)")).IsEqualTo(ErrorValue.Reference);
    }
}
