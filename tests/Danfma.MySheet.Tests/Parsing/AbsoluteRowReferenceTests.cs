using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Whole-row references whose endpoints carry an absolute marker (<c>$1:$1</c>, <c>$1:$1000</c>, the mixed
/// <c>1:$1</c> / <c>$1:1</c>, and their sheet-qualified forms). In Excel the <c>$</c> on a row endpoint is a
/// fill/copy marker only: <c>$1:$1</c> and <c>1:1</c> denote the same range and evaluate identically.
/// Regression for issue #8 (defect 2): the unqualified form evaluated to <c>#REF!</c> and the qualified form
/// threw <c>ParseException</c>, while <c>$A:$A</c> and <c>$A$1:$B$1</c> already worked.
/// </summary>
public class AbsoluteRowReferenceTests
{
    // Row 1 holds three headers, row 2 two numbers — on BOTH sheets, so the same expectations hold for the
    // unqualified (context sheet) and the qualified ("Other Sheet") forms. Formulas are evaluated without
    // being stored in a cell, so a whole-row range never includes the formula itself.
    private static (Workbook Workbook, Sheet Sheet) Sheets()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("S");
        var other = workbook.Sheets.Add("Other Sheet");

        foreach (var target in new[] { sheet, other })
        {
            target["A1"] = String("h1");
            target["B1"] = String("h2");
            target["C1"] = String("h3");
            target["A2"] = Number(10);
            target["B2"] = Number(40);
        }

        return (workbook, sheet);
    }

    private static object? Eval(string formula, Sheet sheet, Workbook workbook) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

    [Test]
    [Arguments("=$1:$1")]
    [Arguments("=1:$1")]
    [Arguments("=$1:1")]
    public async Task Parse_AbsoluteWholeRow_IsTheSameOpenRangeAsTheRelativeForm(string formula)
    {
        var sheet = new Sheet { Name = "S" };

        var expression = ExpressionParser.Parse(formula, sheet);

        await Assert.That(expression).IsEqualTo(new OpenRangeReference(null, null, 1, 1, "S"));
    }

    [Test]
    public async Task Parse_QualifiedAbsoluteWholeRow_IsAnOpenRangeOnThatSheet()
    {
        var sheet = new Sheet { Name = "S" };

        var expression = ExpressionParser.Parse("='Other Sheet'!$1:$1000", sheet);

        await Assert
            .That(expression)
            .IsEqualTo(new OpenRangeReference(null, null, 1, 1000, "Other Sheet"));
    }

    [Test]
    [Arguments("=COUNTA(1:1)", 3.0)] // control: the relative form already worked
    [Arguments("=COUNTA($1:$1)", 3.0)]
    [Arguments("=COUNTA(1:$1)", 3.0)]
    [Arguments("=COUNTA($1:1)", 3.0)]
    [Arguments("=COUNTA($1:$1000)", 5.0)]
    [Arguments("=MATCH(\"h2\",$1:$1,0)", 2.0)]
    [Arguments("=INDEX($1:$1000,2,2)", 40.0)]
    public async Task Unqualified_AbsoluteWholeRow_EvaluatesLikeTheRelativeForm(
        string formula,
        double expected
    )
    {
        var (workbook, sheet) = Sheets();

        await Assert.That(Eval(formula, sheet, workbook) as double?).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=COUNTA('Other Sheet'!1:1)", 3.0)] // control
    [Arguments("=COUNTA('Other Sheet'!$1:$1)", 3.0)]
    [Arguments("=COUNTA('Other Sheet'!1:$1)", 3.0)]
    [Arguments("=COUNTA('Other Sheet'!$1:1)", 3.0)]
    [Arguments("=COUNTA('Other Sheet'!$1:$1000)", 5.0)]
    [Arguments("=MATCH(\"h2\",'Other Sheet'!$1:$1,0)", 2.0)]
    [Arguments("=INDEX('Other Sheet'!$1:$1000,2,2)", 40.0)]
    public async Task Qualified_AbsoluteWholeRow_ParsesAndEvaluates(string formula, double expected)
    {
        var (workbook, sheet) = Sheets();

        await Assert.That(Eval(formula, sheet, workbook) as double?).IsEqualTo(expected);
    }

    [Test]
    public async Task Qualified_AbsoluteWholeRow_InsideLet_MatchesTheIssueFormula()
    {
        // The representative real-world formula from issue #8: a header lookup filled down a column.
        var (workbook, sheet) = Sheets();
        sheet["J1"] = String("h2");
        sheet["B416"] = Number(2);

        var formula =
            "=IF($B416=\"\",\"\",LET(hdr,'Other Sheet'!$1:$1,"
            + "colNum,MATCH(SUBSTITUTE(J$1,\"'\",\"''\"),hdr,0),"
            + "val,INDEX('Other Sheet'!$1:$1000,$B416,colNum),"
            + "IF(TRIM(val&\"\")=\"\",\"\",val)))";

        await Assert.That(Eval(formula, sheet, workbook) as double?).IsEqualTo(40.0);
    }

    [Test]
    public async Task Write_AbsoluteWholeRow_DropsTheMarker_LikeAbsoluteColumns()
    {
        // The AST keeps no '$' for open ranges ($A:$A already writes back as A:A); rows follow suit.
        var sheet = new Sheet { Name = "S" };

        var written = ExpressionParser.Parse("=SUM('Other Sheet'!$1:$1)", sheet).ToFormula("S");

        await Assert.That(written).IsEqualTo("SUM('Other Sheet'!1:1)");
    }

    [Test]
    public async Task BareDollarRow_OutsideARange_IsStillAName()
    {
        // '$1' only means "row 1" as a range endpoint. On its own it stays a (unbound) name → #NAME?,
        // never a number — so `=$1+1` does not silently become 2.
        var (workbook, sheet) = Sheets();

        await Assert.That(Eval("=$1+1", sheet, workbook)).IsEqualTo(ErrorValue.Name);
    }

    // --- The '$<row>' endpoint applies the same ceiling as the numeric-endpoint path (int.MaxValue): a
    // huge row must be rejected, never overflow into a negative or wrapped row number.

    [Test]
    [Arguments("$1", true, 1)]
    [Arguments("1048576", true, 1048576)]
    [Arguments("$2147483647", true, int.MaxValue)]
    [Arguments("$2147483648", false, 0)] // int.MaxValue + 1
    [Arguments("$99999999999", false, 0)]
    [Arguments("$0", false, 0)]
    [Arguments("$", false, 0)]
    [Arguments("$1A", false, 0)]
    public async Task TryParseRow_RejectsZeroNonDigitsAndOverflow(
        string label,
        bool expectedOk,
        int expectedRow
    )
    {
        var ok = CellAddress.TryParseRow(label, out var row);

        await Assert.That(ok).IsEqualTo(expectedOk);
        await Assert.That(row).IsEqualTo(expectedRow);
    }

    [Test]
    public async Task HugeAbsoluteRow_FailsExactlyLikeTheRelativeForm()
    {
        var (workbook, sheet) = Sheets();

        var absolute = Eval("=COUNTA($99999999999:$99999999999)", sheet, workbook);
        var relative = Eval("=COUNTA(99999999999:99999999999)", sheet, workbook);

        await Assert.That(absolute is ErrorValue).IsTrue();
        await Assert.That(absolute).IsEqualTo(relative);
    }
}
