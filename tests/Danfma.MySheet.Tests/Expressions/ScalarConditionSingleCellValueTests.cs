using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Sweep 32/33/34 final-review fix wave, finding C1 (Critical): a single-cell BARE-REFERENCE branch of a
/// scalar-condition <c>IF</c>/<c>CHOOSE</c> must hand a SCALAR consumer the cell's own VALUE — the same
/// convention <c>INDEX(A1:A3,2)+1</c>, <c>OFFSET(A1,0,0)+1</c> and <c>INDIRECT("A1")+1</c> already follow
/// for a producer's or a name's single-cell result. Before this fix (branch <c>5975c63</c>)
/// <c>ArrayEvaluation.BranchValue</c> and <c>Choose.CaptureChosen</c> handed EVERY consumer a
/// <c>ComputedValue.Reference</c>, which no scalar coercion dereferences, so 47 of 93 measured scalar-slot
/// shapes moved from main's (oracle-matching) answer to <c>#VALUE!</c> — the most common IF shape in any
/// workbook, with no test in the suite that reached it (Fable 5.1 final review, section C1).
/// <para>
/// Fixture: the controller's divergence probe grid, Main!A1:A3 = 5, 0, 9, B1:B3 = 1, 2, 3, E2 = <c>=1/0</c>,
/// plus Fable's C1 additions — the defined names <c>CellA1</c> = Main!$A$1 and <c>ErrCell</c> = Main!$E$2, and
/// F1 left blank. The formula sits at Main!AZ5000 and is read back through the cell, mirroring
/// <see cref="LookupValueErrorCellTests"/>'s harness.
/// </para>
/// <para>
/// Every row below is measured on Aspose.Cells 26.7.0, one formula per workbook: PLAIN and CSE agree on
/// every row in this class (Fable 5.1 final review + the controller's own divergence probe, 2026-09-14).
/// </para>
/// </summary>
public class ScalarConditionSingleCellValueTests
{
    private static string InCell(string formula)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(5);
        main["A2"] = new NumberValue(0);
        main["A3"] = new NumberValue(9);
        main["B1"] = new NumberValue(1);
        main["B2"] = new NumberValue(2);
        main["B3"] = new NumberValue(3);
        main["E2"] = ExpressionParser.Parse("=1/0", main);
        workbook.DefineName("CellA1", "Main!$A$1");
        workbook.DefineName("ErrCell", "Main!$E$2");
        main["AZ5000"] = ExpressionParser.Parse(formula, main);

        var value = workbook.GetCellValue("Main", "AZ5000");

        if (value.TryGetError(out var error))
        {
            return error.ToString();
        }

        if (value.TryGetNumber(out var number))
        {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        if (value.TryGetBoolean(out var boolean))
        {
            return boolean ? "TRUE" : "FALSE";
        }

        return value.TryGetText(out var text) ? $"\"{text}\"" : value.Kind.ToString();
    }

    // The regression, straight from a BINARY OPERATOR / comparison over the branch's value: main and the
    // oracle both give 2 / 6 / 6 / 10 / 6 / 6 / TRUE / FALSE / TRUE (branch: #VALUE! on every row before
    // this fix — a Reference-kind value coerces to a number or an equality on neither side).
    [Test]
    [Arguments("=IF(A1>0,B1,C1)*2", "2")]
    [Arguments("=IF(TRUE,A1,0)+1", "6")]
    [Arguments("=IF(FALSE,0,A1)+1", "6")]
    [Arguments("=CHOOSE(2,0,A1)*2", "10")]
    [Arguments("=IF(TRUE,CellA1,0)+1", "6")]
    [Arguments("=CHOOSE(1,CellA1)+1", "6")]
    [Arguments("=IF(TRUE,A1,0)=5", "TRUE")]
    [Arguments("=IF(TRUE,A1,0)<>5", "FALSE")]
    [Arguments("=IF(TRUE,A1,0)=IF(TRUE,A1,0)", "TRUE")]
    public async Task ABinaryOperatorOrComparison_OverTheSingleCellBranch_SeesItsValue(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // ISERROR/IFERROR/NOT/ISNUMBER/ISBLANK/N over the single-cell branch: main and the oracle inspect the
    // cell's OWN value/kind (branch before this fix: FALSE/#DIV/0!/#VALUE! throughout — a Reference-kind
    // value is never an Error, never a Number, never Blank).
    [Test]
    [Arguments("=ISERROR(IF(TRUE,E2,0))", "TRUE")]
    [Arguments("=ISERROR(CHOOSE(1,ErrCell))", "TRUE")]
    [Arguments("=IFERROR(IF(TRUE,E2,0),\"e\")", "\"e\"")]
    [Arguments("=IFERROR(CHOOSE(1,ErrCell),\"e\")", "\"e\"")]
    [Arguments("=IF(ISERROR(IF(TRUE,E2,0)),\"err\",\"ok\")", "\"err\"")]
    [Arguments("=IF(IF(TRUE,A1,0),\"t\",\"f\")", "\"t\"")]
    [Arguments("=NOT(IF(TRUE,A1,0))", "FALSE")]
    [Arguments("=ISNUMBER(IF(TRUE,A1,0))", "TRUE")]
    [Arguments("=ISBLANK(IF(TRUE,F1,0))", "TRUE")]
    [Arguments("=N(IF(TRUE,A1,0))", "5")]
    [Arguments("=IF(TRUE,E2,0)+1", "#DIV/0!")]
    [Arguments("=IF(TRUE,ErrCell,0)+1", "#DIV/0!")]
    public async Task AnInspectionOrErrorHandlingFunction_OverTheSingleCellBranch_SeesItsValue(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // The single-cell branch as a LOOKUP value, a criterion, or a position argument: main and the oracle
    // read the cell's number (branch before this fix: #N/A / 0 / #VALUE! — a Reference-kind value matches
    // nothing in a lookup array, and does not coerce to a criterion or a position).
    [Test]
    [Arguments("=MATCH(IF(TRUE,A1,0),A1:A3,0)", "1")]
    [Arguments("=VLOOKUP(IF(TRUE,A1,0),A1:B3,2,FALSE)", "1")]
    [Arguments("=XLOOKUP(IF(TRUE,A1,0),A1:A3,B1:B3)", "1")]
    [Arguments("=COUNTIF(A1:A3,IF(TRUE,A1,0))", "1")]
    [Arguments("=INDEX(B1:B3,IF(TRUE,B2,0))", "2")]
    [Arguments("=OFFSET(A1,IF(TRUE,B2,0),0)", "9")]
    public async Task ASingleCellBranch_AsALookupValueCriterionOrPosition_IsTheCellsNumber(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Text and math functions over the single-cell branch: main and the oracle read/coerce the cell's own
    // content (branch before this fix: #VALUE! throughout — none of LEN/&/TEXT/REPT/UPPER/ROUND/INT/MOD can
    // coerce a Reference-kind value).
    [Test]
    [Arguments("=LEN(IF(TRUE,A1,\"\"))", "1")]
    [Arguments("=IF(TRUE,A1,0)&\"x\"", "\"5x\"")]
    [Arguments("=TEXT(IF(TRUE,A1,0),\"0\")", "\"5\"")]
    [Arguments("=REPT(\"x\",IF(TRUE,B2,0))", "\"xx\"")]
    [Arguments("=UPPER(IF(TRUE,A2,\"a\"))", "\"0\"")]
    [Arguments("=ROUND(IF(TRUE,A1,0),0)", "5")]
    [Arguments("=ABS(IF(TRUE,A1,0))", "5")]
    [Arguments("=INT(IF(TRUE,A1,0))", "5")]
    [Arguments("=MOD(IF(TRUE,A1,0),2)", "1")]
    public async Task ATextOrMathFunction_OverTheSingleCellBranch_CoercesItsValue(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // A single-cell branch as an argument inside DATE, SEQUENCE, LET and SWITCH — the shapes furthest from
    // a bare scalar slot. DATE(2020,IF(TRUE,B2,0),1) with B2=2 is 2020-02-01, Excel serial 43862.
    [Test]
    [Arguments("=DATE(2020,IF(TRUE,B2,0),1)", "43862")]
    [Arguments("=SUM(SEQUENCE(IF(TRUE,B2,0)))", "3")]
    [Arguments("=LET(x,IF(TRUE,A1,0),x+1)", "6")]
    [Arguments("=SWITCH(IF(TRUE,A1,0),5,\"five\",\"other\")", "\"five\"")]
    public async Task ASingleCellBranch_InsideDateSequenceLetOrSwitch_SeesItsValue(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // The controls: unaffected before and after this fix. A REFERENCE-AWARE consumer (comparison against a
    // literal range's cell here is really the criteria/index family already covered above; this is the
    // plain, ordinary boundary case with no selector at all) and a bare-number single-cell branch that was
    // never the Reference-kind wrapper in the first place.
    [Test]
    [Arguments("=IF(A1>0,B1,C1)", "1")]
    [Arguments("=SUM(IF(TRUE,A1,0))", "5")]
    public async Task TheOrdinaryControls_AreUnmoved(string formula, string expected)
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }
}
