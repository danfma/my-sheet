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
            .That(
                Calc("=XOR(A1:A2)", ("A1", Expression.String("a")), ("A2", Expression.String("b")))
            )
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Xor_PropagatesErrors()
    {
        await Assert.That(Calc("=XOR(TRUE,1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=XOR(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }

    // --- OR / AND / XOR: text and blank inside a reference are ignored (Excel semantics) ---
    // support.microsoft.com AND/OR/XOR remarks: "If a specified range contains no logical values,
    // [the function] returns the #VALUE! error." Text and empty cells reached THROUGH a reference or
    // array argument are ignored; #VALUE! is returned only when no evaluable value survives at all.
    // Regression oracle: K1 comparison against Aspose.Cells 26.6 / Excel (MySheet 2.9.0 gave #VALUE!).

    [Test]
    public async Task Or_IgnoresTextFromReference()
    {
        // =OR(A192="Show", A208) with A192="Show", A208="Hide" (text): the comparison is TRUE and the
        // text cell A208 is ignored -> TRUE (Excel/Aspose). MySheet 2.9.0 returned #VALUE!.
        var cells = new (string, Expression)[]
        {
            ("A192", Expression.String("Show")),
            ("A208", Expression.String("Hide")),
        };

        await Assert.That(Calc("=OR(A192=\"Show\", A208)", cells) as bool?).IsTrue();
    }

    [Test]
    public async Task Or_ReturnsFalseWhenOnlyFalseSurvives()
    {
        // =OR(FALSE, A208): the text cell is ignored, only FALSE remains evaluable -> FALSE.
        await Assert
            .That(Calc("=OR(FALSE, A208)", ("A208", Expression.String("Hide"))) as bool?)
            .IsFalse();
    }

    [Test]
    public async Task Or_SingleTextReferenceIsValueError()
    {
        // =OR(A208) with A208 text: nothing evaluable survives -> #VALUE!.
        await Assert
            .That(Calc("=OR(A208)", ("A208", Expression.String("Hide"))))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task And_IgnoresTextFromReference()
    {
        // =AND(TRUE, A208): the text cell is ignored, only TRUE remains -> TRUE.
        await Assert
            .That(Calc("=AND(TRUE, A208)", ("A208", Expression.String("Hide"))) as bool?)
            .IsTrue();
    }

    [Test]
    public async Task And_SingleTextReferenceIsValueError()
    {
        // =AND(A208) with A208 text: nothing evaluable survives -> #VALUE!.
        await Assert
            .That(Calc("=AND(A208)", ("A208", Expression.String("Hide"))))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Xor_IgnoresTextFromSingleCellReference()
    {
        // The same rule applies to a SINGLE cell reference, not just a range: =XOR(TRUE, A208) with
        // A208 text ignores the text and leaves one TRUE -> TRUE (odd count). This closed a latent
        // XOR gap: the range path already ignored text, but a single CellReference did not.
        await Assert
            .That(Calc("=XOR(TRUE, A208)", ("A208", Expression.String("Hide"))) as bool?)
            .IsTrue();
    }

    [Test]
    public async Task Or_NumberFromReferenceIsEvaluated()
    {
        // A number reached through a reference still evaluates (non-zero -> TRUE), text alongside is
        // ignored: =OR(A1, A2) with A1=5 (number), A2="x" (text) -> TRUE.
        var cells = new (string, Expression)[]
        {
            ("A1", Expression.Number(5)),
            ("A2", Expression.String("x")),
        };

        await Assert.That(Calc("=OR(A1, A2)", cells) as bool?).IsTrue();
    }

    [Test]
    public async Task Or_LiteralTextArgumentIsIgnored()
    {
        // Confirmed against the Aspose/K1 oracle doc (2026-07-03): a LITERAL text argument follows the
        // SAME rule as text reached through a reference — it is IGNORED, not #VALUE!.
        //   =OR(TRUE, "literal text") -> TRUE  (text ignored, TRUE survives)
        //   =OR(FALSE, "x")           -> FALSE (text ignored, FALSE survives)
        //   =OR("x")                  -> #VALUE! (nothing evaluable survives)
        await Assert.That(Calc("=OR(TRUE, \"literal text\")") as bool?).IsTrue();
        await Assert.That(Calc("=OR(FALSE, \"x\")") as bool?).IsFalse();
        await Assert.That(Calc("=OR(\"x\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task And_LiteralTextArgumentIsIgnored()
    {
        // Same family, same reducer: AND ignores a literal text argument.
        await Assert.That(Calc("=AND(TRUE, \"x\")") as bool?).IsTrue();
        await Assert.That(Calc("=AND(FALSE, \"x\")") as bool?).IsFalse();
        await Assert.That(Calc("=AND(\"x\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Xor_LiteralTextArgumentIsIgnored()
    {
        // Same family, same reducer: XOR ignores a literal text argument (odd-TRUE count over the rest).
        await Assert.That(Calc("=XOR(TRUE, \"x\")") as bool?).IsTrue();
        await Assert.That(Calc("=XOR(TRUE, TRUE, \"x\")") as bool?).IsFalse();
        await Assert.That(Calc("=XOR(\"x\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Or_TextFromConcatenationIsIgnored()
    {
        // A scalar whose value is text produced by a comparison/arithmetic/CONCATENATION is text, hence
        // ignored: =OR(FALSE, A1 & "x") with A1="p" yields the text "px" -> ignored -> only FALSE -> FALSE.
        await Assert
            .That(Calc("=OR(FALSE, A1 & \"x\")", ("A1", Expression.String("p"))) as bool?)
            .IsFalse();
    }

    [Test]
    public async Task Or_LiteralNumberAndBooleanStillEvaluate()
    {
        // Text is ignored WITHOUT coercion, but literal numbers (≠0 -> TRUE) and booleans still evaluate,
        // and an ERROR argument still propagates (an error is not ignorable text).
        await Assert.That(Calc("=OR(FALSE, 0, 3)") as bool?).IsTrue(); // 3 ≠ 0 -> TRUE
        await Assert.That(Calc("=AND(TRUE, 0)") as bool?).IsFalse(); // 0 -> FALSE
        await Assert.That(Calc("=OR(FALSE, \"x\", 1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task Or_BlankCellReferenceIsIgnored()
    {
        // A blank cell reached through a reference is ignored like text: =OR(A1) with A1 empty leaves
        // nothing evaluable -> #VALUE! (Excel). This changes the pre-fix blank->FALSE coercion.
        await Assert.That(Calc("=OR(A1)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Or_CrossSheetTextReferenceIsIgnored_K1Case()
    {
        // The exact K1 divergence: H23 = =IF(OR(Sheet8!A192="Show",Sheet8!A208),"*","") with
        // Sheet8!A192="Show" and Sheet8!A208="Hide" (text). Excel/Aspose -> "*"; 2.9.0 gave #VALUE!.
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Sheet1");
        var sheet8 = workbook.Sheets.Add("Sheet8");
        sheet8["A192"] = Expression.String("Show");
        sheet8["A208"] = Expression.String("Hide");

        var result = ExpressionParser
            .Parse("=IF(OR(Sheet8!A192=\"Show\",Sheet8!A208),\"*\",\"\")", main)
            .Evaluate(workbook)
            .AsObject();

        await Assert.That(result as string).IsEqualTo("*");
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
            .That(
                Calc("=SWITCH(99,1,\"Sunday\",2,\"Monday\",3,\"Tuesday\",\"No match\")") as string
            )
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
