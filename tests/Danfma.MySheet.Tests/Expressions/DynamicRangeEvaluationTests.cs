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
        ExpressionParser
            .Parse(formula, sheet)
            .Evaluate(new EvaluationContext(wb, sheet.Name))
            .AsObject();

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
        await Assert
            .That(Calc(wb, sheet, "=COUNTIF(INDEX(A1:A5,2):A5,30)") as double?)
            .IsEqualTo(2.0);
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
        await Assert
            .That(Calc(wb, sheet, "=SUM(INDEX(ROW($A:$A),2):A1)"))
            .IsEqualTo(ErrorValue.Reference);
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
    // reference/NameReference argument. BOTH endpoints are on Sheet2 (which does not exist — Grid() adds
    // only "Sheet1"), so this exercises the MISSING-sheet guard, not the cross-sheet guard. The dynamic
    // range resolves structurally to a range on the missing sheet; COUNT must short-circuit to #REF!
    // rather than silently enumerating zero cells and returning 0.
    [Test]
    public async Task CountFamily_DynamicRangeOverMissingSheet_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));
        await Assert
            .That(Calc(wb, sheet, "=COUNT(INDEX(Sheet2!A1:A3,2):Sheet2!A5)"))
            .IsEqualTo(ErrorValue.Reference);
    }

    // Cross-sheet ':' is #REF! in Excel: Sheet2 EXISTS here, so this is not a missing-sheet case — the
    // endpoints simply live on different sheets (INDEX target on Sheet2, A5 on Sheet1), which Excel rejects.
    [Test]
    public async Task DynamicRange_CrossSheetEndpoints_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));
        var other = wb.Sheets.Add("Sheet2");
        other["A1"] = new NumberValue(9);
        other["A2"] = new NumberValue(9);
        other["A3"] = new NumberValue(9);

        await Assert
            .That(Calc(wb, sheet, "=COUNT(INDEX(Sheet2!A1:A3,2):A5)"))
            .IsEqualTo(ErrorValue.Reference);
        // Same-sheet control still resolves (both endpoints on Sheet1).
        await Assert.That(Calc(wb, sheet, "=SUM(INDEX(A1:A3,2):A3)") as double?).IsEqualTo(5.0);
    }

    // Static both-qualified cross-sheet range Sheet1!A1:Sheet2!B2 is #REF! in Excel (both sheets exist
    // here). Previously the parser threw; now it builds a DynamicRange that resolves cross-sheet → #REF!.
    [Test]
    public async Task StaticCrossSheetQualifiedRange_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 1));
        wb.Sheets.Add("Sheet2")["B2"] = new NumberValue(9);
        await Assert
            .That(Calc(wb, sheet, "=SUM(Sheet1!A1:Sheet2!B2)"))
            .IsEqualTo(ErrorValue.Reference);
    }

    // Static both-qualified SAME-sheet range Sheet1!A1:Sheet1!A3 is a valid range (previously threw).
    [Test]
    public async Task StaticSameSheetQualifiedRange_Resolves()
    {
        var (wb, sheet) = Grid(("A1", 1), ("A2", 2), ("A3", 3));
        await Assert.That(Calc(wb, sheet, "=SUM(Sheet1!A1:Sheet1!A3)") as double?).IsEqualTo(6.0);
    }

    // Aspose oracle: Excel truncates OFFSET height/width toward zero, so OFFSET(A1,0,0,1.9,1) is a
    // SINGLE cell (A1), not a 2-row range. SUM == A1 == 10, NOT A1+A2 == 30.
    [Test]
    public async Task Offset_FractionalHeightWidth_TruncatesLikeExcel()
    {
        var (wb, sheet) = Grid(("A1", 10), ("A2", 20));
        await Assert.That(Calc(wb, sheet, "=SUM(OFFSET(A1,0,0,1.9,1))") as double?).IsEqualTo(10.0);
    }

    // Aspose oracle: a non-positive OFFSET size is #REF! in Excel (height/width of 0, or a fraction that
    // truncates to 0). Guards against building an invalid "A0" cell id from a zero-size range.
    [Test]
    public async Task Offset_NonPositiveHeightWidth_IsRefError()
    {
        var (wb, sheet) = Grid(("A1", 10), ("A2", 20));
        await Assert
            .That(Calc(wb, sheet, "=SUM(OFFSET(A1,0,0,0,1))"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc(wb, sheet, "=SUM(OFFSET(A1,0,0,0.9,1))"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc(wb, sheet, "=SUM(OFFSET(A1,0,0,1,0))"))
            .IsEqualTo(ErrorValue.Reference);
    }
}
