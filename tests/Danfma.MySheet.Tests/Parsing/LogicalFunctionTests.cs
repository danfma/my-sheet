using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// TRUE/FALSE/XOR/IFS/SWITCH. Oracles: the official Microsoft support pages for each function
/// (fetched 2026-07-01), cited per test.
/// </summary>
public class LogicalFunctionTests
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
    public async Task True_And_False_AreZeroArgFunctions()
    {
        // support.microsoft.com TRUE/FALSE: no arguments; provided for compatibility.
        await Assert.That(Calc("=TRUE()") as bool?).IsTrue();
        await Assert.That(Calc("=FALSE()") as bool?).IsFalse();
        await Assert.That(Calc("=IF(1=1,TRUE())") as bool?).IsTrue();
    }

    [Test]
    public async Task Xor_MatchesExcelDocs()
    {
        // support.microsoft.com XOR: XOR(3>0,2<9)=FALSE (both TRUE); XOR(3>12,4>6)=FALSE (both FALSE).
        await Assert.That(Calc("=XOR(3>0,2<9)") as bool?).IsFalse();
        await Assert.That(Calc("=XOR(3>12,4>6)") as bool?).IsFalse();
    }

    [Test]
    public async Task Xor_IsTrueForOddNumberOfTrueInputs()
    {
        // support.microsoft.com XOR remarks: "the result of XOR is TRUE when the number of TRUE
        // inputs is odd and FALSE when the number of TRUE inputs is even."
        await Assert.That(Calc("=XOR(TRUE)") as bool?).IsTrue();
        await Assert.That(Calc("=XOR(TRUE,FALSE)") as bool?).IsTrue();
        await Assert.That(Calc("=XOR(TRUE,TRUE,TRUE)") as bool?).IsTrue();
        await Assert.That(Calc("=XOR(TRUE,TRUE,TRUE,TRUE)") as bool?).IsFalse();
        await Assert.That(Calc("=XOR(FALSE)") as bool?).IsFalse();
    }

    [Test]
    public async Task Xor_RangeIgnoresTextAndBlanks()
    {
        // support.microsoft.com XOR remarks: text and empty cells in a reference argument are
        // ignored; a range with no logical values at all -> #VALUE!.
        await Assert
            .That(
                Calc(
                    "=XOR(A1:A4)",
                    ("A1", Expression.String("skip")),
                    ("A2", new BooleanValue(true)),
                    ("A3", Expression.Number(0)),
                    ("A4", new BooleanValue(true))
                ) as bool?
            )
            .IsFalse(); // TRUE + 0(FALSE) + TRUE = two TRUEs -> even -> FALSE

        await Assert
            .That(Calc("=XOR(A1:A2)", ("A1", Expression.String("a")), ("A2", Expression.String("b"))))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Xor_PropagatesErrors()
    {
        await Assert.That(Calc("=XOR(TRUE,1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=XOR(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Ifs_ReturnsFirstMatch()
    {
        // support.microsoft.com IFS grades example: =IFS(A2>89,"A",A2>79,"B",A2>69,"C",A2>59,"D",
        // TRUE,"F") — 93 -> "A"; 58 falls through to the TRUE default -> "F" (rows 6-7 of the doc).
        const string grades = "=IFS(A2>89,\"A\",A2>79,\"B\",A2>69,\"C\",A2>59,\"D\",TRUE,\"F\")";

        await Assert.That(Calc(grades, ("A2", Expression.Number(93))) as string).IsEqualTo("A");
        await Assert.That(Calc(grades, ("A2", Expression.Number(85))) as string).IsEqualTo("B");
        await Assert.That(Calc(grades, ("A2", Expression.Number(58))) as string).IsEqualTo("F");
    }

    [Test]
    public async Task Ifs_NoTrueConditionIsNA()
    {
        // support.microsoft.com IFS remarks: "If no TRUE conditions are found, this function
        // returns #N/A error."
        await Assert.That(Calc("=IFS(1>2,\"a\",3>4,\"b\")")).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task Ifs_IsLazy()
    {
        // Only the matched pair's value is computed, and conditions after the match are skipped —
        // same short-circuit contract as IF.
        await Assert.That(Calc("=IFS(TRUE,1,TRUE,1/0)") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=IFS(FALSE,1/0,TRUE,2)") as double?).IsEqualTo(2.0);
        await Assert.That(Calc("=IFS(TRUE,1,1/0,2)") as double?).IsEqualTo(1.0);
    }

    [Test]
    public async Task Ifs_ConditionErrors()
    {
        // support.microsoft.com IFS remarks: a logical_test resolving to a non-logical -> #VALUE!.
        await Assert.That(Calc("=IFS(\"abc\",1)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=IFS(1/0,1)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Switch_MatchesExcelDocs()
    {
        // support.microsoft.com SWITCH example table: SWITCH(99,1,"Sunday",2,"Monday",3,"Tuesday")
        // = #N/A; with "No match" default = "No match"; SWITCH(2,1,"Sunday",7,"Saturday","weekday")
        // = "weekday"; SWITCH(3,1,"Sunday",2,"Monday",3,"Tuesday","No match") = "Tuesday".
        await Assert
            .That(Calc("=SWITCH(99,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\")"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Calc("=SWITCH(99,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\",\"No match\")") as string)
            .IsEqualTo("No match");
        await Assert
            .That(Calc("=SWITCH(2,1,\"Sunday\",7,\"Saturday\",\"weekday\")") as string)
            .IsEqualTo("weekday");
        await Assert
            .That(Calc("=SWITCH(3,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\",\"No match\")") as string)
            .IsEqualTo("Tuesday");
    }

    [Test]
    public async Task Switch_ComparesLikeEquality()
    {
        // Equality follows the '=' operator semantics (ValueCoercion.AreEqual): text matches
        // case-insensitively, and a number never equals its text form.
        await Assert.That(Calc("=SWITCH(\"AB\",\"ab\",1,2)") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=SWITCH(1,\"1\",\"text\",\"none\")") as string).IsEqualTo("none");
    }

    [Test]
    public async Task Switch_IsLazy()
    {
        // The matched result short-circuits: later value/result pairs are never computed.
        await Assert.That(Calc("=SWITCH(1,1,\"one\",1/0,\"boom\")") as string).IsEqualTo("one");
        await Assert.That(Calc("=SWITCH(1,1,\"one\",2,1/0)") as string).IsEqualTo("one");
    }

    [Test]
    public async Task Switch_PropagatesErrors()
    {
        await Assert.That(Calc("=SWITCH(1/0,1,\"one\")")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SWITCH(1,1/0,\"one\",2)")).IsEqualTo(ErrorValue.DivByZero);
    }
}
