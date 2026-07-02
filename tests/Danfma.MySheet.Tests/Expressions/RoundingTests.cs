using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// ROUNDDOWN/TRUNC/MROUND/CEILING*/FLOOR*/EVEN/ODD. The sign semantics of the legacy CEILING/FLOOR
/// and the MROUND/EVEN/ODD values are golden values from the Microsoft support pages (fetched
/// 2026-07-01); the *.PRECISE variants follow their definition (round toward ±infinity, significance
/// sign ignored), which is also ISO.CEILING per ODF OpenFormula.
/// </summary>
public class RoundingTests
{
    private const double Tolerance = 1e-9;

    private static object? Calc(string formula, params (string Id, Expression Value)[] cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        foreach (var (id, value) in cells)
        {
            sheet[id] = value;
        }

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    [Test]
    public async Task RoundDown_TruncatesTowardZero()
    {
        await Assert.That(Num(Calc("=ROUNDDOWN(3.14159,3)"))).IsEqualTo(3.141).Within(Tolerance);
        await Assert.That(Num(Calc("=ROUNDDOWN(-3.7,0)"))).IsEqualTo(-3.0);
        await Assert.That(Num(Calc("=ROUNDDOWN(12345,-2)"))).IsEqualTo(12300.0);
        // Binary-noise guard: 2.3*10 is 22.999999999999996 in IEEE-754; Excel still yields 2.3.
        await Assert.That(Num(Calc("=ROUNDDOWN(2.3,1)"))).IsEqualTo(2.3).Within(Tolerance);
    }

    [Test]
    public async Task Trunc_MatchesExcelDocs()
    {
        // support.microsoft.com TRUNC: TRUNC(8.9)=8; TRUNC(-8.9)=-8; TRUNC(0.45)=0.
        await Assert.That(Num(Calc("=TRUNC(8.9)"))).IsEqualTo(8.0);
        await Assert.That(Num(Calc("=TRUNC(-8.9)"))).IsEqualTo(-8.0);
        await Assert.That(Num(Calc("=TRUNC(0.45)"))).IsEqualTo(0.0);
        await Assert.That(Num(Calc("=TRUNC(3.14159,2)"))).IsEqualTo(3.14).Within(Tolerance);
        await Assert.That(Num(Calc("=TRUNC(2.3,1)"))).IsEqualTo(2.3).Within(Tolerance);
    }

    [Test]
    public async Task MRound_MatchesExcelDocs()
    {
        // support.microsoft.com MROUND: MROUND(10,3)=9; MROUND(-10,-3)=-9; MROUND(1.3,0.2)=1.4;
        // MROUND(5,-2)=#NUM! (arguments must share the sign).
        await Assert.That(Num(Calc("=MROUND(10,3)"))).IsEqualTo(9.0);
        await Assert.That(Num(Calc("=MROUND(-10,-3)"))).IsEqualTo(-9.0);
        await Assert.That(Num(Calc("=MROUND(1.3,0.2)"))).IsEqualTo(1.4).Within(Tolerance);
        await Assert.That(Calc("=MROUND(5,-2)")).IsEqualTo(ErrorValue.Number);
    }

    [Test]
    public async Task Ceiling_MatchesExcelDocs()
    {
        // support.microsoft.com CEILING: CEILING(2.5,1)=3; CEILING(-2.5,-2)=-4 (both negative →
        // away from zero); CEILING(-2.5,2)=-2 (rounded up toward zero); CEILING(1.5,0.1)=1.5;
        // CEILING(0.234,0.01)=0.24.
        await Assert.That(Num(Calc("=CEILING(2.5,1)"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=CEILING(-2.5,-2)"))).IsEqualTo(-4.0);
        await Assert.That(Num(Calc("=CEILING(-2.5,2)"))).IsEqualTo(-2.0);
        await Assert.That(Num(Calc("=CEILING(1.5,0.1)"))).IsEqualTo(1.5).Within(Tolerance);
        await Assert.That(Num(Calc("=CEILING(0.234,0.01)"))).IsEqualTo(0.24).Within(Tolerance);
    }

    [Test]
    public async Task Ceiling_ErrorsAndZeroSignificance()
    {
        // Positive number with negative significance → #NUM! (legacy CEILING, mirror of the rule
        // documented on the FLOOR page). Significance 0 → 0 (Excel/ODF OpenFormula behaviour).
        await Assert.That(Calc("=CEILING(2.5,-2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Num(Calc("=CEILING(2.5,0)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task CeilingMath_MatchesExcelDocs()
    {
        // support.microsoft.com CEILING.MATH: CEILING.MATH(24.3,5)=25; CEILING.MATH(6.7)=7;
        // CEILING.MATH(-8.1,2)=-8 (default rounds negatives toward zero);
        // CEILING.MATH(-5.5,2,-1)=-6 (non-zero mode rounds negatives away from zero).
        await Assert.That(Num(Calc("=CEILING.MATH(24.3,5)"))).IsEqualTo(25.0);
        await Assert.That(Num(Calc("=CEILING.MATH(6.7)"))).IsEqualTo(7.0);
        await Assert.That(Num(Calc("=CEILING.MATH(-8.1,2)"))).IsEqualTo(-8.0);
        await Assert.That(Num(Calc("=CEILING.MATH(-5.5,2,-1)"))).IsEqualTo(-6.0);
        // Docs remark: "The mode argument does not affect positive numbers."
        await Assert.That(Num(Calc("=CEILING.MATH(6.3,1,1)"))).IsEqualTo(7.0);
        // Docs remark: mode 1 with significance 1 rounds -6.3 away from zero, to -7.
        await Assert.That(Num(Calc("=CEILING.MATH(-6.3,1,1)"))).IsEqualTo(-7.0);
    }

    [Test]
    public async Task CeilingPrecise_And_IsoCeiling_RoundTowardPositiveInfinity()
    {
        // By definition (and ODF OpenFormula), the significance sign is ignored and rounding is
        // always toward +infinity.
        await Assert.That(Num(Calc("=CEILING.PRECISE(4.3)"))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=CEILING.PRECISE(-4.3)"))).IsEqualTo(-4.0);
        await Assert.That(Num(Calc("=CEILING.PRECISE(4.3,-2)"))).IsEqualTo(6.0);
        await Assert.That(Num(Calc("=CEILING.PRECISE(-4.3,2)"))).IsEqualTo(-4.0);
        await Assert.That(Num(Calc("=ISO.CEILING(4.3)"))).IsEqualTo(5.0);
        await Assert.That(Num(Calc("=ISO.CEILING(-4.3,-2)"))).IsEqualTo(-4.0);
    }

    [Test]
    public async Task Floor_MatchesExcelDocs()
    {
        // support.microsoft.com FLOOR: FLOOR(3.7,2)=2; FLOOR(-2.5,-2)=-2 (both negative → toward
        // zero); FLOOR(2.5,-2)=#NUM!; FLOOR(1.58,0.1)=1.5; FLOOR(0.234,0.01)=0.23.
        await Assert.That(Num(Calc("=FLOOR(3.7,2)"))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=FLOOR(-2.5,-2)"))).IsEqualTo(-2.0);
        await Assert.That(Calc("=FLOOR(2.5,-2)")).IsEqualTo(ErrorValue.Number);
        await Assert.That(Num(Calc("=FLOOR(1.58,0.1)"))).IsEqualTo(1.5).Within(Tolerance);
        await Assert.That(Num(Calc("=FLOOR(0.234,0.01)"))).IsEqualTo(0.23).Within(Tolerance);
        // Negative number with positive significance rounds down, away from zero.
        await Assert.That(Num(Calc("=FLOOR(-2.5,2)"))).IsEqualTo(-4.0);
    }

    [Test]
    public async Task Floor_ZeroSignificanceIsDivZero()
    {
        // Legacy FLOOR with significance 0 → #DIV/0! (ECMA-376 / Excel behaviour; CEILING
        // asymmetrically yields 0).
        await Assert.That(Calc("=FLOOR(2.5,0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task FloorMath_MatchesExcelDocs()
    {
        // support.microsoft.com FLOOR.MATH: FLOOR.MATH(24.3,5)=20; FLOOR.MATH(6.7)=6;
        // FLOOR.MATH(-8.1,2)=-10 (default rounds negatives away from zero);
        // FLOOR.MATH(-5.5,2,-1)=-4 (non-zero mode rounds negatives toward zero).
        await Assert.That(Num(Calc("=FLOOR.MATH(24.3,5)"))).IsEqualTo(20.0);
        await Assert.That(Num(Calc("=FLOOR.MATH(6.7)"))).IsEqualTo(6.0);
        await Assert.That(Num(Calc("=FLOOR.MATH(-8.1,2)"))).IsEqualTo(-10.0);
        await Assert.That(Num(Calc("=FLOOR.MATH(-5.5,2,-1)"))).IsEqualTo(-4.0);
        // Mode does not affect positive numbers.
        await Assert.That(Num(Calc("=FLOOR.MATH(6.7,1,1)"))).IsEqualTo(6.0);
    }

    [Test]
    public async Task FloorPrecise_RoundsTowardNegativeInfinity()
    {
        await Assert.That(Num(Calc("=FLOOR.PRECISE(4.3)"))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=FLOOR.PRECISE(-4.3)"))).IsEqualTo(-5.0);
        await Assert.That(Num(Calc("=FLOOR.PRECISE(-4.3,-2)"))).IsEqualTo(-6.0);
    }

    [Test]
    public async Task Even_MatchesExcelDocs()
    {
        // support.microsoft.com EVEN: EVEN(1.5)=2; EVEN(3)=4; EVEN(2)=2; EVEN(-1)=-2
        // (away from zero; an even integer is unchanged).
        await Assert.That(Num(Calc("=EVEN(1.5)"))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=EVEN(3)"))).IsEqualTo(4.0);
        await Assert.That(Num(Calc("=EVEN(2)"))).IsEqualTo(2.0);
        await Assert.That(Num(Calc("=EVEN(-1)"))).IsEqualTo(-2.0);
        await Assert.That(Num(Calc("=EVEN(0)"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Odd_MatchesExcelDocs()
    {
        // support.microsoft.com ODD: ODD(1.5)=3; ODD(3)=3; ODD(2)=3; ODD(-1)=-1; ODD(-2)=-3.
        // ODD(0)=1 (zero is even, so it rounds away from zero to the next odd).
        await Assert.That(Num(Calc("=ODD(1.5)"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ODD(3)"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ODD(2)"))).IsEqualTo(3.0);
        await Assert.That(Num(Calc("=ODD(-1)"))).IsEqualTo(-1.0);
        await Assert.That(Num(Calc("=ODD(-2)"))).IsEqualTo(-3.0);
        await Assert.That(Num(Calc("=ODD(0)"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task ArgumentErrors_Propagate()
    {
        await Assert.That(Calc("=CEILING(1/0,1)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=TRUNC(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=EVEN(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }
}
