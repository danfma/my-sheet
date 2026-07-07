using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Task 7: end-to-end evaluation of the ':' range operator over reference-returning expressions
/// (<see cref="DynamicRange"/>), through the production consumers (SUM/COUNTIF), plus the Excel-compat
/// and hardening cases from the design's self-review. Expected numeric values for the SUM/COUNTIF/corner
/// cases were verified against Aspose.Cells' <c>CalculateFormula()</c> (see the design's Aspose oracle
/// procedure); all three matched hand computation exactly (90, 2, 10).
/// </summary>
public class DynamicRangeEvaluationTests
{
    private static (Workbook, Sheet) Grid(params (string Id, double Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        foreach (var (id, value) in cells)
            sheet[id] = new NumberValue(value);
        return (workbook, sheet);
    }

    private static object? Calc(Workbook wb, Sheet sheet, string formula) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(new EvaluationContext(wb, sheet.Name)).AsObject();

    // Aspose oracle: SUM over the dynamic range INDEX(A1:A5,2):A4 == A2+A3+A4 == 20+30+40 == 90.
    [Test]
    public async Task Sum_OverDynamicRange_MatchesExcel()
    {
        var (wb, sheet) = Grid(("A1", 10), ("A2", 20), ("A3", 30), ("A4", 40), ("A5", 50));
        await Assert.That(Calc(wb, sheet, "=SUM(INDEX(A1:A5,2):A4)") as double?).IsEqualTo(90.0);
    }

    // Aspose oracle: COUNTIF(INDEX(A1:A5,2):A5, 30) spans A2:A5 = [10,30,40,30]; A3 and A5 match -> 2.
    [Test]
    public async Task CountIf_OverDynamicRange_MatchesExcel()
    {
        var (wb, sheet) = Grid(("A1", 30), ("A2", 10), ("A3", 30), ("A4", 40), ("A5", 30));
        await Assert.That(Calc(wb, sheet, "=COUNTIF(INDEX(A1:A5,2):A5,30)") as double?).IsEqualTo(2.0);
    }

    // Corner rule: reversed endpoints normalise (OFFSET(B2,0,0):A1 behaves like B2:A1 == A1:B2).
    // Aspose oracle: sum of the 2x2 block A1+B1+A2+B2 == 1+2+3+4 == 10.
    [Test]
    public async Task ReversedCorners_Normalise()
    {
        var (wb, sheet) = Grid(("A1", 1), ("B1", 2), ("A2", 3), ("B2", 4));
        await Assert.That(Calc(wb, sheet, "=SUM(OFFSET(B2,0,0):A1)") as double?).IsEqualTo(10.0);
    }

    // Array-form INDEX as an endpoint has no cell address (INDEX(ROW($A:$A),2) is the mini-CSE identity
    // vector, not a reference) -> the whole dynamic range fails to resolve -> #REF! (Excel parity).
    [Test]
    public async Task ArrayIndexEndpoint_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 1));
        await Assert.That(Calc(wb, sheet, "=SUM(INDEX(ROW($A:$A),2):A1)")).IsEqualTo(ErrorValue.Reference);
    }

    // Hardening: a ':' endpoint that is not reference-returning at all (a string literal) must also
    // degrade to #REF!, not throw or silently coerce.
    [Test]
    public async Task NonReferenceEndpoint_IsRefError()
    {
        var (wb, sheet) = Grid();
        await Assert.That(Calc(wb, sheet, "=\"text\":A1")).IsEqualTo(ErrorValue.Reference);
    }

    // Regression (K1 hot idiom): INDEX(ROW($A:$A),k) must still return the row number through the VALUE
    // path (not be redirected into reference resolution) even though INDEX/ROW now also support
    // TryResolveReference for other shapes.
    [Test]
    public async Task IndexIntoRowArray_StillReturnsRowNumber()
    {
        var (wb, sheet) = Grid();
        await Assert.That(Calc(wb, sheet, "=INDEX(ROW($A:$A),3)") as double?).IsEqualTo(3.0);
    }

    // Regression: ReferenceGuard.MissingSheet must guard a DynamicRange too, same as it does a plain
    // reference/NameReference argument. Sheet2 does not exist in this workbook (only "Sheet1" is added by
    // Grid()) — the dynamic range INDEX(Sheet2!A1:A3,2):A5 resolves structurally (INDEX/DynamicRange don't
    // themselves check sheet existence) to a range on the missing sheet. Before the fix, COUNT's
    // error-ignoring fold silently enumerated zero cells and returned 0; it must short-circuit to #REF!.
    [Test]
    public async Task CountFamily_DynamicRangeOverMissingSheet_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));
        await Assert.That(Calc(wb, sheet, "=COUNT(INDEX(Sheet2!A1:A3,2):A5)")).IsEqualTo(ErrorValue.Reference);
    }
}
