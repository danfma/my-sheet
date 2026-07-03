using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using static Danfma.MySheet.Expressions.Expression;

namespace Danfma.MySheet.Tests;

/// <summary>
/// Excel-parity guard for empty formula results: at the CELL boundary (<see cref="Workbook.GetCellValue"/>)
/// a cell that HAS content whose formula evaluates to blank displays 0, while a truly empty cell stays
/// blank. Internal expression semantics (blank compares as ""/0/FALSE inside an expression) are unchanged,
/// which these tests also pin.
/// </summary>
public class EmptyFormulaResultParityTests
{
    [Test]
    public async Task DirectReferenceToEmptyCell_CoercesToZero()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // A1 = "=F10"; F10 is empty.
        sheet["A1"] = Cell("F10", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0.0);
    }

    [Test]
    public async Task CrossSheetReferenceToEmptyCell_CoercesToZero()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        workbook.Sheets.Add("Sheet2");

        // Sheet1!A1 = "=Sheet2!F10"; Sheet2!F10 is empty.
        sheet1["A1"] = Cell("F10", "Sheet2");

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0.0);
    }

    [Test]
    public async Task IfBranchReturningEmptyCell_CoercesToZero()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // A1 = "=IF(TRUE, F10)"; F10 is empty.
        sheet["A1"] = ExpressionParser.Parse("=IF(TRUE, F10)", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0.0);
    }

    [Test]
    public async Task TrulyEmptyCell_StaysBlank()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add("Sheet1");

        // Z99 was never assigned: no content, no formula.
        var value = workbook.GetCellValue("Sheet1", "Z99");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Blank);
    }

    [Test]
    public async Task ExplicitBlankValueCell_StaysBlank()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // A cell explicitly holding the empty expression is still "truly empty".
        sheet["A1"] = BlankValue.Instance;

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Blank);
    }

    [Test]
    public async Task IsBlank_TrulyEmptyCell_IsTrue()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // A1 = "=ISBLANK(F10)"; F10 truly empty.
        sheet["A1"] = ExpressionParser.Parse("=ISBLANK(F10)", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(value.ToBoolean()).IsTrue();
    }

    [Test]
    public async Task IsBlank_FormulaEmptyCell_IsFalse()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // B1 = "=F10" (formula-empty, now worth 0). A1 = "=ISBLANK(B1)" → FALSE, because B1 is no longer
        // blank at the cell boundary. This is Excel's behavior.
        sheet["B1"] = Cell("F10", sheet);
        sheet["A1"] = ExpressionParser.Parse("=ISBLANK(B1)", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Boolean);
        await Assert.That(value.ToBoolean()).IsFalse();
    }

    [Test]
    public async Task Count_OverFormulaEmptyCell_CountsIt()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // B1 = "=F10" (formula-empty → 0, a number). COUNT now sees a number.
        sheet["B1"] = Cell("F10", sheet);
        sheet["A1"] = ExpressionParser.Parse("=COUNT(B1)", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(1.0);
    }

    [Test]
    public async Task CountBlank_OverFormulaEmptyCell_DoesNotCountIt()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // B1 = "=F10" (formula-empty → 0). COUNTBLANK no longer treats it as blank.
        sheet["B1"] = Cell("F10", sheet);
        sheet["A1"] = ExpressionParser.Parse("=COUNTBLANK(B1)", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(0.0);
    }

    [Test]
    public async Task InternalBlankComparison_IsPreserved()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // =IF(F10="",1,2) with F10 truly empty → 1: INTERNALLY, blank still equals "" inside the expression.
        // (F10 is read via GetCellValue but is a truly empty cell, so it is NOT coerced.)
        sheet["A1"] = ExpressionParser.Parse("=IF(F10=\"\",1,2)", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(value.ToDouble()).IsEqualTo(1.0);
    }

    [Test]
    public async Task ConcatenationWithEmptyCell_StaysEmptyString()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        // =F10&"" with F10 empty → "" : the cell result is Text (not blank), so no coercion; the internal
        // blank→"" concatenation is preserved.
        sheet["A1"] = ExpressionParser.Parse("=F10&\"\"", sheet);

        var value = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(value.Kind).IsEqualTo(ComputedValueKind.Text);
        await Assert.That(value.ToText()).IsEqualTo("");
    }

    [Test]
    public async Task CoercedZero_IsCached_SecondReadIdentical()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Cell("F10", sheet);

        var first = workbook.GetCellValue("Sheet1", "A1");
        var second = workbook.GetCellValue("Sheet1", "A1");

        await Assert.That(first.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(first.ToDouble()).IsEqualTo(0.0);
        await Assert.That(second.Kind).IsEqualTo(ComputedValueKind.Number);
        await Assert.That(second.ToDouble()).IsEqualTo(0.0);
    }
}
