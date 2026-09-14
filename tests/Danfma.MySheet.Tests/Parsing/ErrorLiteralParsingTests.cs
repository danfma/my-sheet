using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// Item 43 (sweep 31-35-43): an Excel error literal (<c>#REF!</c>, <c>#N/A</c>, …) inside formula text.
/// Excel writes a broken reference back as <c>#REF!</c>, so a real <c>.xlsx</c> holding one used to
/// degrade through <see cref="Danfma.MySheet.Excel.ExcelLoadWarningKind.UnparsableFormula"/> — the
/// tokenizer threw <c>ParseException</c> before the parser ever saw the literal. Measured on
/// Aspose.Cells 26.7.0 with 26.6.0 beside it (2026-09-14, byte-identical on every row here): the 7
/// classic error literals parse as ordinary formula text and propagate exactly like any other error
/// value already does. See <c>.superpowers/sdd/sweep-31-35-43/progress.md</c> (Phase 1 checkpoint A)
/// for the full oracle table.
/// </summary>
public class ErrorLiteralParsingTests
{
    private static (Workbook Workbook, Sheet Sheet) NewWorkbook(string sheetName = "Sheet1")
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add(sheetName);

        return (workbook, sheet);
    }

    private static object? Calc(string formula, Sheet sheet, Workbook workbook) =>
        ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

    [Test]
    public async Task BareErrorLiteral_EvaluatesToItself()
    {
        var (workbook, sheet) = NewWorkbook();

        await Assert.That(Calc("=#NULL!", sheet, workbook)).IsEqualTo(ErrorValue.Null);
        await Assert.That(Calc("=#DIV/0!", sheet, workbook)).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=#VALUE!", sheet, workbook)).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=#REF!", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=#NAME?", sheet, workbook)).IsEqualTo(ErrorValue.Name);
        await Assert.That(Calc("=#NUM!", sheet, workbook)).IsEqualTo(ErrorValue.Number);
        await Assert.That(Calc("=#N/A", sheet, workbook)).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task ErrorLiteral_IsCaseInsensitive()
    {
        var (workbook, sheet) = NewWorkbook();

        await Assert.That(Calc("=#ref!", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=#Name?", sheet, workbook)).IsEqualTo(ErrorValue.Name);
        await Assert.That(Calc("=#n/a", sheet, workbook)).IsEqualTo(ErrorValue.NotAvailable);
    }

    // The brief's seven acceptance formulas (Aspose PLAIN column), confirmed on the oracle before any
    // code was written — see the ledger checkpoint. #REF! propagates through every consumer exactly
    // like it already does when produced at evaluation time (a ghost-sheet reference, a deleted range).
    [Test]
    public async Task TheBriefsAcceptanceFormulas_MatchAsposePlain()
    {
        var (workbook, sheet) = NewWorkbook();

        await Assert.That(Calc("=#REF!", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=SUM(#REF!)", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=IF(ISNA(#N/A),1,0)", sheet, workbook) as double?).IsEqualTo(1.0);
        await Assert
            .That(Calc("=MATCH(#REF!,#REF!,0)", sheet, workbook))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=MATCH(1,#REF!,0)", sheet, workbook))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=IFNA(MATCH(#REF!,#REF!,0),\"na\")", sheet, workbook))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=COUNTIF(#REF!,#REF!)", sheet, workbook))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task ErrorLiteral_UnderAnOperator()
    {
        var (workbook, sheet) = NewWorkbook();

        await Assert.That(Calc("=#REF!+1", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=1+#REF!", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=-#N/A", sheet, workbook)).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=#REF!*2", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=#REF!&\"x\"", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task ErrorLiteral_InAComparison()
    {
        var (workbook, sheet) = NewWorkbook();
        sheet["A1"] = new NumberValue(1);

        await Assert.That(Calc("=A1=#N/A", sheet, workbook)).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=#REF!=#N/A", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task SheetQualified_ErrorLiteral_EvaluatesToTheError()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        workbook.Sheets.Add("My Sheet");

        await Assert.That(Calc("=Sheet1!#REF!", sheet1, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("='My Sheet'!#REF!", sheet1, workbook))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=SUM(Sheet1!#REF!)", sheet1, workbook))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task RangeEndpoint_ReferenceErrorLiteral_EvaluatesToReference()
    {
        var (workbook, sheet) = NewWorkbook();
        sheet["A1"] = new NumberValue(5);

        await Assert.That(Calc("=A1:#REF!", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=#REF!:A1", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
        await Assert.That(Calc("=SUM(A1:#REF!)", sheet, workbook)).IsEqualTo(ErrorValue.Reference);
    }

    // Round 2, I-1: measured (Aspose.Cells 26.7.0/26.6.0, PLAIN=CSE, 2026-09-14), a sheet qualifier's
    // '!' accepts the SAME single exception as a ':' range endpoint — only #REF!. Every other error
    // literal there is "Invalid data before reference sign" on the oracle; `Sheet1!#REF!` itself keeps
    // working (SheetQualified_ErrorLiteral_EvaluatesToTheError above).
    [Test]
    [Arguments("=Sheet1!#N/A")]
    [Arguments("='My Sheet'!#DIV/0!")]
    [Arguments("=SUM(Sheet1!#VALUE!)")]
    public async Task SheetQualified_NonReferenceErrorLiteral_IsAParseException(string formula)
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        workbook.Sheets.Add("My Sheet");

        var exception = Assert.Throws<ParseException>(() =>
            ExpressionParser.Parse(formula, sheet1)
        );

        await Assert.That(exception.Kind).IsEqualTo(ParseErrorKind.ExpectedCellReference);
    }

    // Round 2, I-3: Excel writes a DELETED SHEET's qualifier as #REF! itself, e.g. `=#REF!A1` for what
    // used to be `=Other!A1`. Measured (Aspose.Cells 26.7.0/26.6.0, PLAIN=CSE, 2026-09-14): #REF! in
    // this PREFIX position (not after a real sheet name) consumes the reference-shaped text glued
    // directly after it — a cell, a range, another #REF!, a redundant '!', or a parenthesized group —
    // and the whole run evaluates to #REF!, same as the qualifier case. An ordinary operator or
    // terminator (+, *, ^, %, =, &, a function's comma, a closing paren) is never absorbed: it applies
    // normally to the resulting #REF! value.
    [Test]
    [Arguments("=#REF!A1")]
    [Arguments("=SUM(#REF!A1:A3)")]
    [Arguments("=#REF!#REF!")]
    [Arguments("=#REF!!A1")]
    [Arguments("=#REF!(1)")]
    [Arguments("=#REF!A1+1")]
    [Arguments("=#REF!A1:A3")] // bare range, not inside a function — absorbed at the top level too
    [Arguments("=#REF!A1*2")]
    [Arguments("=(#REF!A1)")]
    [Arguments("=IF(TRUE,#REF!A1,0)")]
    [Arguments("=#REF!A1=5")]
    [Arguments("=#REF!A1&\"x\"")]
    [Arguments("=-#REF!A1")]
    [Arguments("=#REF!A1%")]
    [Arguments("=#REF!A1^2")]
    [Arguments("=SUM(#REF!A1,1)")]
    [Arguments("=#REF!A1:#REF!")] // both endpoints corrupted
    public async Task DeletedSheetQualifier_ConsumesTheFollowingReference_EvaluatesToReference(
        string formula
    )
    {
        var (workbook, sheet) = NewWorkbook();
        sheet["A1"] = new NumberValue(5);

        await Assert.That(Calc(formula, sheet, workbook)).IsEqualTo(ErrorValue.Reference);
    }

    // Round 2, I-3 near misses: this absorption is specific to #REF! (the one literal Excel itself
    // writes as a qualifier); a bare non-#REF! literal glued to what follows — with or without a space
    // — is still "Unrecognized" on the oracle, i.e. a genuine trailing token neither the tokenizer nor
    // the parser can attach to the formula. Already true before round 2 (the tokenizer reads exactly
    // one literal and stops; ParseFormula rejects a dangling token) — pinned so it stays true.
    [Test]
    [Arguments("=#N/A A1")]
    [Arguments("=#N/AA1")]
    public async Task NonReferenceErrorLiteral_GluedToAReference_IsStillAParseException(
        string formula
    )
    {
        var (_, sheet) = NewWorkbook();

        await Assert.That(() => ExpressionParser.Parse(formula, sheet)).Throws<ParseException>();
    }

    // Round 2, M-1: the qualified ':' branch (Sheet1!A1:B2) had no arm for an Error token at either
    // endpoint, so only the LEFT-open form worked (Sheet1!#REF!:A1, via ParseQualifiedReference's own
    // early return before the ':' is even looked at). Measured (Aspose.Cells 26.7.0/26.6.0,
    // PLAIN=CSE): both directions are #REF!.
    [Test]
    public async Task SheetQualified_RangeEndpoint_ReferenceErrorLiteral_EvaluatesToReference()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        sheet1["A1"] = new NumberValue(5);

        await Assert
            .That(Calc("=Sheet1!A1:#REF!", sheet1, workbook))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(Calc("=Sheet1!#REF!:A1", sheet1, workbook))
            .IsEqualTo(ErrorValue.Reference);
    }

    // Measured discrimination (2026-09-14, both Aspose versions): ONLY #REF! is accepted as a range
    // endpoint. Any other error literal there is a parse error on the oracle ("Invalid data
    // after/before range sign ':'"), and was already a ParseException in MySheet before this item too
    // (for a different reason — the tokenizer didn't recognise '#' at all). The fix keeps it a
    // ParseException rather than silently accepting a shape the oracle rejects.
    [Test]
    [Arguments("=A1:#N/A")]
    [Arguments("=#N/A:A1")]
    [Arguments("=A1:#DIV/0!")]
    [Arguments("=#VALUE!:A1")]
    public async Task RangeEndpoint_NonReferenceErrorLiteral_IsAParseException(string formula)
    {
        var (_, sheet) = NewWorkbook();

        await Assert.That(() => ExpressionParser.Parse(formula, sheet)).Throws<ParseException>();
    }

    // The brief's explicit round-trip pin: FORMULATEXT of a cell containing =SUM(#REF!) must read the
    // literal back, "=" included — proving the parser produced a real tree (not a degrade) AND the
    // writer renders ErrorValue inside a function argument correctly.
    [Test]
    public async Task FormulaText_OverACellHoldingAnErrorLiteral_RendersTheLiteralBack()
    {
        var (workbook, sheet) = NewWorkbook();
        sheet["A1"] = ExpressionParser.Parse("=SUM(#REF!)", sheet);

        await Assert
            .That(Calc("=FORMULATEXT(A1)", sheet, workbook) as string)
            .IsEqualTo("=SUM(#REF!)");
    }
}
