using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// REGEXTEST/REGEXEXTRACT/REGEXREPLACE. Oracles: the official Microsoft support pages (fetched
/// 2026-07-01 — regextest-function-7d38200b, regexextract-function-4b96c140,
/// regexreplace-function-9c030bb2), cited per test. Excel specifies the PCRE2 flavor; the engine
/// uses System.Text.RegularExpressions (documented limitation) with a defensive 1s timeout.
/// </summary>
public class RegexFunctionTests
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
    public async Task RegexTest_MatchesExcelDocs()
    {
        // support.microsoft.com REGEXTEST example 1 (A2="alfalfa"): "a"=TRUE; "[a-z]"=TRUE;
        // "[A-Z]"=FALSE; "[aeiou]"=TRUE; "[0-9]"=FALSE.
        var a2 = ("A2", Expression.String("alfalfa"));

        await Assert.That(Calc("=REGEXTEST(A2,\"a\")", a2) as bool?).IsTrue();
        await Assert.That(Calc("=REGEXTEST(A2,\"[a-z]\")", a2) as bool?).IsTrue();
        await Assert.That(Calc("=REGEXTEST(A2,\"[A-Z]\")", a2) as bool?).IsFalse();
        await Assert.That(Calc("=REGEXTEST(A2,\"[aeiou]\")", a2) as bool?).IsTrue();
        await Assert.That(Calc("=REGEXTEST(A2,\"[0-9]\")", a2) as bool?).IsFalse();
    }

    [Test]
    public async Task RegexTest_PhonePatternAndCaseSensitivity()
    {
        // support.microsoft.com REGEXTEST example 2: pattern ^\([0-9]{3}\) [0-9]{3}-[0-9]{4}$ —
        // "(378) 555-4195" -> TRUE; "+1(878) 555-8622" -> FALSE. case_sensitivity: 0 (default)
        // sensitive, 1 insensitive.
        const string pattern = "\"^\\([0-9]{3}\\) [0-9]{3}-[0-9]{4}$\"";

        await Assert
            .That(Calc($"=REGEXTEST(\"(378) 555-4195\",{pattern})") as bool?)
            .IsTrue();
        await Assert
            .That(Calc($"=REGEXTEST(\"+1(878) 555-8622\",{pattern})") as bool?)
            .IsFalse();
        await Assert.That(Calc("=REGEXTEST(\"ABC\",\"[a-z]+\",1)") as bool?).IsTrue();
        await Assert.That(Calc("=REGEXTEST(\"ABC\",\"[a-z]+\",0)") as bool?).IsFalse();
    }

    [Test]
    public async Task RegexTest_InvalidInputs()
    {
        // Invalid pattern or an unsupported case_sensitivity flag -> #VALUE! (the pages document
        // no error rows; this follows Excel behavior).
        await Assert.That(Calc("=REGEXTEST(\"a\",\"[\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=REGEXTEST(\"a\",\"a\",2)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task RegexExtract_MatchesExcelDocs()
    {
        // support.microsoft.com REGEXEXTRACT example 1 (A2="DylanWilliams"):
        // REGEXEXTRACT(A2,"[A-Z][a-z]+") = "Dylan" (return_mode 0 = first match, the default).
        var a2 = ("A2", Expression.String("DylanWilliams"));

        await Assert.That(Calc("=REGEXEXTRACT(A2,\"[A-Z][a-z]+\")", a2) as string).IsEqualTo("Dylan");
        await Assert
            .That(Calc("=REGEXEXTRACT(A2,\"[A-Z][a-z]+\",0)", a2) as string)
            .IsEqualTo("Dylan");
        // A phone number out of the doc's example-2 data (scalar first-match form).
        await Assert
            .That(Calc("=REGEXEXTRACT(\"Sonia Rees (378) 555-4195\",\"[0-9()]+ [0-9-]+\")") as string)
            .IsEqualTo("(378) 555-4195");
    }

    [Test]
    public async Task RegexExtract_UnsupportedModesAndNoMatch()
    {
        // return_mode 1 (all matches) and 2 (capturing groups) return arrays — deferred to the
        // arrays phase (F2), so the engine reports #VALUE! for them. No match -> #N/A.
        await Assert.That(Calc("=REGEXEXTRACT(\"ab\",\"a\",1)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=REGEXEXTRACT(\"ab\",\"a\",2)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=REGEXEXTRACT(\"abc\",\"[0-9]\")")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=REGEXEXTRACT(\"a\",\"[\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task RegexReplace_MatchesExcelDocs()
    {
        // support.microsoft.com REGEXREPLACE example 2 (A2="SoniaBrown"):
        // REGEXREPLACE(A2,"([A-Z][a-z]+)([A-Z][a-z]+)","$2, $1") = "Brown, Sonia".
        // Example 1 (single line of the illustration): digits+dash masked with ***-.
        await Assert
            .That(
                Calc("=REGEXREPLACE(\"SoniaBrown\",\"([A-Z][a-z]+)([A-Z][a-z]+)\",\"$2, $1\")")
                    as string
            )
            .IsEqualTo("Brown, Sonia");
        await Assert
            .That(Calc("=REGEXREPLACE(\"Sonia Rees(378) 555-4195\",\"[0-9]+-\",\"***-\")") as string)
            .IsEqualTo("Sonia Rees(378) ***-4195");
    }

    [Test]
    public async Task RegexReplace_OccurrenceModes()
    {
        // support.microsoft.com REGEXREPLACE: occurrence 0 (default) replaces all instances; a
        // positive n replaces only the nth; a negative n counts from the end.
        await Assert.That(Calc("=REGEXREPLACE(\"a1b2c3\",\"[0-9]\",\"*\")") as string).IsEqualTo("a*b*c*");
        await Assert
            .That(Calc("=REGEXREPLACE(\"a1b2c3\",\"[0-9]\",\"*\",2)") as string)
            .IsEqualTo("a1b*c3");
        await Assert
            .That(Calc("=REGEXREPLACE(\"a1b2c3\",\"[0-9]\",\"*\",-1)") as string)
            .IsEqualTo("a1b2c*");
        await Assert.That(Calc("=REGEXREPLACE(\"a\",\"[\",\"x\")")).IsEqualTo(ErrorValue.NotValue);
    }
}
