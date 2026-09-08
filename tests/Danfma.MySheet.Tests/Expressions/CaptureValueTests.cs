using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// <see cref="NamedReferences.CaptureValue"/> is the single rule behind "a range node bound to a name / chosen
/// by CHOOSE / passed through unary + stays a range". These tests cover the arms the LET tests do not: CHOOSE
/// over a union, and the shared-formula <see cref="AnchoredRangeReference"/> arm, which must resolve to the
/// rectangle of THE SLAVE being evaluated, not the master's.
/// </summary>
public class CaptureValueTests
{
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");
        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Number(3);
        sheet["A4"] = Number(4);

        return (workbook, sheet);
    }

    [Test]
    public async Task Choose_Union_StaysARange()
    {
        var (workbook, sheet) = Grid();

        var value = ExpressionParser
            .Parse("=SUM(CHOOSE(1,(A1:A1,A3:A4)))", sheet)
            .Evaluate(workbook);

        await Assert.That(value.AsDouble()).IsEqualTo(8.0);
    }

    [Test]
    [Arguments("SUM(CHOOSE(1,A1:A2))")]
    [Arguments("LET(r,A1:A2,SUM(r))")]
    [Arguments("SUM(+A1:A2)")]
    public async Task AnchoredRange_InSharedFormula_ResolvesPerSlave(string body)
    {
        // Master at row 1 reads A1:A2 (=3); the slave one row down must read A2:A3 (=5), i.e. the anchored
        // range is shifted by the slave's delta before being captured — not frozen at the master's rectangle.
        var (workbook, sheet) = Grid();
        var master = ExpressionParser.ParseAnchoredMasterBody(
            ExpressionParser.TokenizeFormulaBody(body),
            sheet
        );

        var atMaster = new SharedFormulaSlave(master, 0, 0).Evaluate(workbook);
        var oneRowDown = new SharedFormulaSlave(master, 1, 0).Evaluate(workbook);

        await Assert.That(atMaster.AsDouble()).IsEqualTo(3.0);
        await Assert.That(oneRowDown.AsDouble()).IsEqualTo(5.0);
    }
}
