using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// NA, the IS* family, N, T, TYPE, ERROR.TYPE and SHEETS. Oracles: the official Microsoft support
/// pages ("IS functions", ISEVEN, ISODD, ISFORMULA, NA, N, T, TYPE, ERROR.TYPE, SHEETS — fetched
/// 2026-07-01), cited per test.
/// </summary>
public class InformationFunctionTests
{
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

    [Test]
    public async Task Na_ReturnsNotAvailable()
    {
        // support.microsoft.com NA: "Returns the error value #N/A"; no arguments.
        await Assert.That(Calc("=NA()")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=ISNA(NA())") as bool?).IsTrue();
    }

    [Test]
    public async Task IsError_IsErr_IsNa_MatchExcelDocs()
    {
        // support.microsoft.com "IS functions" example 2 (A4=#REF!, A6=#N/A):
        // ISERROR(#REF!)=TRUE; ISNA(#REF!)=FALSE; ISNA(#N/A)=TRUE; ISERR(#N/A)=FALSE.
        await Assert.That(Calc("=ISERROR(A4)", ("A4", ErrorValue.Reference)) as bool?).IsTrue();
        await Assert.That(Calc("=ISNA(A4)", ("A4", ErrorValue.Reference)) as bool?).IsFalse();
        await Assert.That(Calc("=ISNA(A6)", ("A6", ErrorValue.NotAvailable)) as bool?).IsTrue();
        await Assert.That(Calc("=ISERR(A6)", ("A6", ErrorValue.NotAvailable)) as bool?).IsFalse();

        // ISERR is TRUE for every error except #N/A; non-errors are FALSE for all three.
        await Assert.That(Calc("=ISERR(1/0)") as bool?).IsTrue();
        await Assert.That(Calc("=ISERROR(1)") as bool?).IsFalse();
        await Assert.That(Calc("=ISNA(\"x\")") as bool?).IsFalse();
    }

    [Test]
    public async Task IsText_IsNonText_IsLogical_DoNotCoerce()
    {
        // support.microsoft.com "IS functions": ISLOGICAL(TRUE)=TRUE; ISLOGICAL("TRUE")=FALSE
        // (values are NOT converted); ISTEXT("Region1")=TRUE; ISNONTEXT is TRUE for blanks.
        await Assert.That(Calc("=ISLOGICAL(TRUE)") as bool?).IsTrue();
        await Assert.That(Calc("=ISLOGICAL(\"TRUE\")") as bool?).IsFalse();
        await Assert
            .That(Calc("=ISTEXT(A3)", ("A3", Expression.String("Region1"))) as bool?)
            .IsTrue();
        await Assert.That(Calc("=ISTEXT(19)") as bool?).IsFalse();
        await Assert.That(Calc("=ISNONTEXT(19)") as bool?).IsTrue();
        await Assert.That(Calc("=ISNONTEXT(\"x\")") as bool?).IsFalse();
        await Assert.That(Calc("=ISNONTEXT(A9)") as bool?).IsTrue(); // blank cell
    }

    [Test]
    public async Task IsEven_MatchesExcelDocs()
    {
        // support.microsoft.com ISEVEN: ISEVEN(-1)=FALSE; ISEVEN(2.5)=TRUE (truncated to 2);
        // ISEVEN(5)=FALSE; ISEVEN(0)=TRUE; ISEVEN over serial 40900 = TRUE; nonnumeric -> #VALUE!.
        await Assert.That(Calc("=ISEVEN(-1)") as bool?).IsFalse();
        await Assert.That(Calc("=ISEVEN(2.5)") as bool?).IsTrue();
        await Assert.That(Calc("=ISEVEN(5)") as bool?).IsFalse();
        await Assert.That(Calc("=ISEVEN(0)") as bool?).IsTrue();
        await Assert.That(Calc("=ISEVEN(40900)") as bool?).IsTrue();
        await Assert.That(Calc("=ISEVEN(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task IsOdd_MatchesExcelDocs()
    {
        // support.microsoft.com ISODD: ISODD(-1)=TRUE; ISODD(2.5)=FALSE (truncated to 2);
        // ISODD(5)=TRUE; nonnumeric -> #VALUE!.
        await Assert.That(Calc("=ISODD(-1)") as bool?).IsTrue();
        await Assert.That(Calc("=ISODD(2.5)") as bool?).IsFalse();
        await Assert.That(Calc("=ISODD(5)") as bool?).IsTrue();
        await Assert.That(Calc("=ISODD(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task IsRef_IsASyntacticCheck()
    {
        // support.microsoft.com "IS functions": ISREF(G8)=TRUE. The check is on the argument node
        // itself (cell/range/union reference), regardless of the referenced value.
        await Assert.That(Calc("=ISREF(G8)") as bool?).IsTrue();
        await Assert.That(Calc("=ISREF(A1:A3)") as bool?).IsTrue();
        await Assert.That(Calc("=ISREF((A1:A3,C1))") as bool?).IsTrue();
        await Assert.That(Calc("=ISREF(123)") as bool?).IsFalse();
        await Assert.That(Calc("=ISREF(\"A1\")") as bool?).IsFalse();
        await Assert.That(Calc("=ISREF(1+1)") as bool?).IsFalse();
    }

    [Test]
    public async Task IsFormula_ChecksTheReferencedCellExpression()
    {
        // support.microsoft.com ISFORMULA: a cell containing a formula -> TRUE (even =3/0, whose
        // value is an error); a literal number or text -> FALSE; non-reference argument -> #VALUE!.
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A2"] = ExpressionParser.Parse("=3/0", sheet);
        sheet["A3"] = Expression.Number(7);
        sheet["A4"] = Expression.String("Hello, world!");

        object? Calc(string formula) =>
            ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();

        await Assert.That(Calc("=ISFORMULA(A2)") as bool?).IsTrue();
        await Assert.That(Calc("=ISFORMULA(A3)") as bool?).IsFalse();
        await Assert.That(Calc("=ISFORMULA(A4)") as bool?).IsFalse();
        await Assert.That(Calc("=ISFORMULA(A9)") as bool?).IsFalse(); // empty cell
        await Assert.That(Calc("=ISFORMULA(7)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task N_MatchesExcelDocs()
    {
        // support.microsoft.com N conversion table: number -> itself; TRUE -> 1; FALSE -> 0; an
        // error -> the error; anything else (text, including "7") -> 0.
        await Assert.That(Calc("=N(A2)", ("A2", Expression.Number(7))) as double?).IsEqualTo(7.0);
        await Assert
            .That(Calc("=N(A3)", ("A3", Expression.String("Even"))) as double?)
            .IsEqualTo(0.0);
        await Assert.That(Calc("=N(A4)", ("A4", new BooleanValue(true))) as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=N(FALSE)") as double?).IsEqualTo(0.0);
        await Assert.That(Calc("=N(\"7\")") as double?).IsEqualTo(0.0);
        await Assert.That(Calc("=N(1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task T_MatchesExcelDocs()
    {
        // support.microsoft.com T: text -> the text ("Rainfall"); number and logical -> "" (empty
        // text). Errors propagate (the engine-wide rule; the page's table only covers non-errors).
        await Assert
            .That(Calc("=T(A2)", ("A2", Expression.String("Rainfall"))) as string)
            .IsEqualTo("Rainfall");
        await Assert.That(Calc("=T(A3)", ("A3", Expression.Number(19))) as string).IsEqualTo("");
        await Assert.That(Calc("=T(A4)", ("A4", new BooleanValue(true))) as string).IsEqualTo("");
        await Assert.That(Calc("=T(1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Type_MatchesExcelDocs()
    {
        // support.microsoft.com TYPE: number=1, text=2, logical=4, error=16 — TYPE(A2)=2 for text
        // "Smith"; TYPE("Mr. "&A2)=2; TYPE(2+A2)=16 (the error is inspected, not propagated).
        await Assert
            .That(Calc("=TYPE(A2)", ("A2", Expression.String("Smith"))) as double?)
            .IsEqualTo(2.0);
        await Assert
            .That(Calc("=TYPE(\"Mr. \"&A2)", ("A2", Expression.String("Smith"))) as double?)
            .IsEqualTo(2.0);
        await Assert
            .That(Calc("=TYPE(2+A2)", ("A2", Expression.String("Smith"))) as double?)
            .IsEqualTo(16.0);
        await Assert.That(Calc("=TYPE(1)") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=TYPE(TRUE)") as double?).IsEqualTo(4.0);
        // An empty cell evaluates as 0, so its TYPE is 1 (number) — matching Excel.
        await Assert.That(Calc("=TYPE(A9)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task ErrorType_MapsTheSevenClassicErrors()
    {
        // support.microsoft.com ERROR.TYPE mapping table: #NULL!=1, #DIV/0!=2, #VALUE!=3, #REF!=4,
        // #NAME?=5, #NUM!=6, #N/A=7; anything that is not an error -> #N/A.
        await Assert
            .That(Calc("=ERROR.TYPE(A2)", ("A2", new ErrorValue("#NULL!"))) as double?)
            .IsEqualTo(1.0);
        await Assert.That(Calc("=ERROR.TYPE(1/0)") as double?).IsEqualTo(2.0);
        await Assert
            .That(Calc("=ERROR.TYPE(A2)", ("A2", ErrorValue.NotValue)) as double?)
            .IsEqualTo(3.0);
        await Assert
            .That(Calc("=ERROR.TYPE(A2)", ("A2", ErrorValue.Reference)) as double?)
            .IsEqualTo(4.0);
        await Assert
            .That(Calc("=ERROR.TYPE(A2)", ("A2", ErrorValue.Name)) as double?)
            .IsEqualTo(5.0);
        await Assert
            .That(Calc("=ERROR.TYPE(A2)", ("A2", ErrorValue.Number)) as double?)
            .IsEqualTo(6.0);
        await Assert.That(Calc("=ERROR.TYPE(NA())") as double?).IsEqualTo(7.0);
        await Assert.That(Calc("=ERROR.TYPE(10)")).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Sheets_CountsTheWorkbookSheets()
    {
        // support.microsoft.com SHEETS: with no reference argument, the total number of sheets in
        // the workbook (the doc example returns 3 for a three-sheet workbook).
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        workbook.Sheets.Add("Sheet2");
        workbook.Sheets.Add("Sheet3");

        var result = ExpressionParser.Parse("=SHEETS()", sheet).Evaluate(workbook).AsObject();

        await Assert.That(result as double?).IsEqualTo(3.0);
    }
}
