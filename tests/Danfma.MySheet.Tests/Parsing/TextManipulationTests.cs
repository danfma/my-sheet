using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// RIGHT/FIND/SEARCH/REPLACE/SUBSTITUTE/REPT/PROPER/EXACT/CHAR/CODE/UNICHAR/UNICODE/CLEAN.
/// Oracles: the official Microsoft support pages for each function (fetched 2026-07-01; the FIND
/// page is retired on the live site — golden values come from the archived copy of the same URL),
/// cited per test. Cases the pages leave undocumented (negative counts, SUBSTITUTE case/instance
/// bounds, CHAR range error) follow well-known Excel behavior and are flagged as such.
/// </summary>
public class TextManipulationTests
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
    public async Task Right_MatchesExcelDocs()
    {
        // support.microsoft.com RIGHT: RIGHT("Sale Price",5)="Price"; RIGHT("Stock Number")="r"
        // (num_chars defaults to 1); num_chars beyond the length returns all of text.
        await Assert.That(Calc("=RIGHT(\"Sale Price\",5)") as string).IsEqualTo("Price");
        await Assert.That(Calc("=RIGHT(\"Stock Number\")") as string).IsEqualTo("r");
        await Assert.That(Calc("=RIGHT(\"ab\",5)") as string).IsEqualTo("ab");
        // num_chars must be >= 0 (doc); negative -> #VALUE! (mirror of LEFT).
        await Assert.That(Calc("=RIGHT(\"ab\",-1)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Find_IsCaseSensitiveWithoutWildcards()
    {
        // support.microsoft.com FIND (A2="Miriam McGovern"): FIND("M",A2)=1; FIND("m",A2)=6;
        // FIND("M",A2,3)=8. FIND is case sensitive and takes no wildcards.
        await Assert.That(Calc("=FIND(\"M\",\"Miriam McGovern\")") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=FIND(\"m\",\"Miriam McGovern\")") as double?).IsEqualTo(6.0);
        await Assert.That(Calc("=FIND(\"M\",\"Miriam McGovern\",3)") as double?).IsEqualTo(8.0);

        // Wildcard characters are literal in FIND.
        await Assert.That(Calc("=FIND(\"?\",\"a?b\")") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Find_ErrorCases()
    {
        // support.microsoft.com FIND remarks: not found -> #VALUE!; start_num not > 0 -> #VALUE!;
        // start_num beyond the length -> #VALUE!; empty find_text matches at start_num.
        await Assert.That(Calc("=FIND(\"x\",\"abc\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=FIND(\"a\",\"abc\",0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=FIND(\"a\",\"abc\",4)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=FIND(\"\",\"abc\")") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=FIND(\"\",\"abc\",2)") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Search_MatchesExcelDocs()
    {
        // support.microsoft.com SEARCH (A2="Statements", A3="Profit Margin", A4="margin",
        // A5='The "boss" is here.'): SEARCH("e",A2,6)=7; SEARCH(A4,A3)=8 (case-insensitive);
        // SEARCH("""",A5)=5.
        await Assert.That(Calc("=SEARCH(\"e\",\"Statements\",6)") as double?).IsEqualTo(7.0);
        await Assert.That(Calc("=SEARCH(\"margin\",\"Profit Margin\")") as double?).IsEqualTo(8.0);
        await Assert
            .That(Calc("=SEARCH(\"\"\"\",\"The \"\"boss\"\" is here.\")") as double?)
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task Search_SupportsWildcards()
    {
        // support.microsoft.com SEARCH remarks: "?" matches any single character, "*" any sequence
        // of characters, and "~" before ?/* searches the literal character. The match is
        // positional: the result is where the pattern starts inside within_text.
        await Assert.That(Calc("=SEARCH(\"p?rt\",\"Report\")") as double?).IsEqualTo(3.0);
        await Assert.That(Calc("=SEARCH(\"m*n\",\"Profit Margin\")") as double?).IsEqualTo(8.0);
        await Assert.That(Calc("=SEARCH(\"*in\",\"Margin\")") as double?).IsEqualTo(1.0);
        await Assert.That(Calc("=SEARCH(\"~?\",\"What? Yes\")") as double?).IsEqualTo(5.0);
        await Assert.That(Calc("=SEARCH(\"~*\",\"2*3\")") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task Search_ErrorCases()
    {
        // support.microsoft.com SEARCH remarks: not found -> #VALUE!; start_num not > 0 or beyond
        // the length of within_text -> #VALUE!.
        await Assert.That(Calc("=SEARCH(\"z\",\"abc\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SEARCH(\"a\",\"abc\",0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=SEARCH(\"a\",\"abc\",9)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Replace_MatchesExcelDocs()
    {
        // support.microsoft.com REPLACE: REPLACE("abcdefghijk",6,5,"*")="abcde*k";
        // REPLACE("2009",3,2,"10")="2010"; REPLACE("123456",1,3,"@")="@456".
        await Assert
            .That(Calc("=REPLACE(\"abcdefghijk\",6,5,\"*\")") as string)
            .IsEqualTo("abcde*k");
        await Assert.That(Calc("=REPLACE(\"2009\",3,2,\"10\")") as string).IsEqualTo("2010");
        await Assert.That(Calc("=REPLACE(\"123456\",1,3,\"@\")") as string).IsEqualTo("@456");
        // Positions are 1-based; start_num < 1 or negative num_chars -> #VALUE! (Excel behavior;
        // the page documents no error rows).
        await Assert.That(Calc("=REPLACE(\"abc\",0,1,\"x\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=REPLACE(\"abc\",1,-1,\"x\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Substitute_MatchesExcelDocs()
    {
        // support.microsoft.com SUBSTITUTE: SUBSTITUTE("Sales Data","Sales","Cost")="Cost Data";
        // SUBSTITUTE("Quarter 1, 2008","1","2",1)="Quarter 2, 2008" (first instance only);
        // SUBSTITUTE("Quarter 1, 2011","1","2",3)="Quarter 1, 2012" (third instance only).
        await Assert
            .That(Calc("=SUBSTITUTE(\"Sales Data\",\"Sales\",\"Cost\")") as string)
            .IsEqualTo("Cost Data");
        await Assert
            .That(Calc("=SUBSTITUTE(\"Quarter 1, 2008\",\"1\",\"2\",1)") as string)
            .IsEqualTo("Quarter 2, 2008");
        await Assert
            .That(Calc("=SUBSTITUTE(\"Quarter 1, 2011\",\"1\",\"2\",3)") as string)
            .IsEqualTo("Quarter 1, 2012");
    }

    [Test]
    public async Task Substitute_IsCaseSensitiveAndValidatesInstance()
    {
        // SUBSTITUTE matches old_text case-sensitively, and instance_num must be >= 1 (well-known
        // Excel behavior; the support page documents neither). An instance beyond the occurrences,
        // or an empty old_text, leaves the text unchanged.
        await Assert
            .That(Calc("=SUBSTITUTE(\"Sales sales\",\"sales\",\"cost\")") as string)
            .IsEqualTo("Sales cost");
        await Assert.That(Calc("=SUBSTITUTE(\"aaa\",\"a\",\"b\",5)") as string).IsEqualTo("aaa");
        await Assert.That(Calc("=SUBSTITUTE(\"abc\",\"\",\"x\")") as string).IsEqualTo("abc");
        await Assert
            .That(Calc("=SUBSTITUTE(\"abc\",\"a\",\"x\",0)"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Rept_MatchesExcelDocs()
    {
        // support.microsoft.com REPT: REPT("*-",3)="*-*-*-"; REPT("-",10)="----------";
        // number_times 0 -> ""; non-integer counts are truncated; a result longer than 32,767
        // characters -> #VALUE!. Negative counts -> #VALUE! (Excel behavior; not on the page).
        await Assert.That(Calc("=REPT(\"*-\",3)") as string).IsEqualTo("*-*-*-");
        await Assert.That(Calc("=REPT(\"-\",10)") as string).IsEqualTo("----------");
        await Assert.That(Calc("=REPT(\"x\",0)") as string).IsEqualTo("");
        await Assert.That(Calc("=REPT(\"ab\",2.9)") as string).IsEqualTo("abab");
        await Assert.That(Calc("=REPT(\"x\",-1)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=REPT(\"ab\",20000)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Proper_MatchesExcelDocs()
    {
        // support.microsoft.com PROPER: "this is a TITLE" -> "This Is A Title";
        // "2-way street" -> "2-Way Street"; "76BudGet" -> "76Budget".
        await Assert
            .That(Calc("=PROPER(\"this is a TITLE\")") as string)
            .IsEqualTo("This Is A Title");
        await Assert.That(Calc("=PROPER(\"2-way street\")") as string).IsEqualTo("2-Way Street");
        await Assert.That(Calc("=PROPER(\"76BudGet\")") as string).IsEqualTo("76Budget");
    }

    [Test]
    public async Task Exact_MatchesExcelDocs()
    {
        // support.microsoft.com EXACT: "word"/"word" -> TRUE; "Word"/"word" -> FALSE (case
        // matters); "w ord"/"word" -> FALSE.
        await Assert.That(Calc("=EXACT(\"word\",\"word\")") as bool?).IsTrue();
        await Assert.That(Calc("=EXACT(\"Word\",\"word\")") as bool?).IsFalse();
        await Assert.That(Calc("=EXACT(\"w ord\",\"word\")") as bool?).IsFalse();
    }

    [Test]
    public async Task Char_And_Code_MatchExcelDocs()
    {
        // support.microsoft.com CHAR: CHAR(65)="A"; CHAR(33)="!"; valid range is 1-255 (out of
        // range -> #VALUE!, Excel behavior). CODE: CODE("A")=65; CODE("!")=33; empty -> #VALUE!.
        await Assert.That(Calc("=CHAR(65)") as string).IsEqualTo("A");
        await Assert.That(Calc("=CHAR(33)") as string).IsEqualTo("!");
        await Assert.That(Calc("=CHAR(0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=CHAR(256)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=CODE(\"A\")") as double?).IsEqualTo(65.0);
        await Assert.That(Calc("=CODE(\"!\")") as double?).IsEqualTo(33.0);
        await Assert.That(Calc("=CODE(\"Alpha\")") as double?).IsEqualTo(65.0); // first character
        await Assert.That(Calc("=CODE(\"\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task UniChar_And_Unicode_MatchExcelDocs()
    {
        // support.microsoft.com UNICHAR: UNICHAR(66)="B"; UNICHAR(32)=" "; UNICHAR(0) -> #VALUE!;
        // out-of-range -> #VALUE!; partial surrogates -> #N/A. UNICODE(" ")=32; UNICODE("B")=66.
        await Assert.That(Calc("=UNICHAR(66)") as string).IsEqualTo("B");
        await Assert.That(Calc("=UNICHAR(32)") as string).IsEqualTo(" ");
        await Assert.That(Calc("=UNICHAR(960)") as string).IsEqualTo("π"); // full Unicode
        await Assert.That(Calc("=UNICHAR(0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=UNICHAR(1114112)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=UNICHAR(55296)")).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=UNICODE(\" \")") as double?).IsEqualTo(32.0);
        await Assert.That(Calc("=UNICODE(\"B\")") as double?).IsEqualTo(66.0);
        await Assert.That(Calc("=UNICODE(\"\")")).IsEqualTo(ErrorValue.NotValue);
        // Astral characters resolve to the full code point (surrogate pair read as one).
        await Assert.That(Calc("=UNICODE(UNICHAR(128169))") as double?).IsEqualTo(128169.0);
    }

    [Test]
    public async Task Clean_RemovesControlCharacters()
    {
        // support.microsoft.com CLEAN: removes characters 0-31 — CLEAN(CHAR(9)&"Monthly report"&
        // CHAR(10)) = "Monthly report". Characters >= 32 (including 127) are kept.
        await Assert
            .That(Calc("=CLEAN(CHAR(9)&\"Monthly report\"&CHAR(10))") as string)
            .IsEqualTo("Monthly report");
        await Assert.That(Calc("=CLEAN(\"abc\")") as string).IsEqualTo("abc");
        await Assert.That(Calc("=LEN(CLEAN(UNICHAR(127)&\"a\"))") as double?).IsEqualTo(2.0);
    }

    [Test]
    public async Task TextManipulation_PropagatesErrors()
    {
        await Assert.That(Calc("=RIGHT(1/0,1)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=FIND(\"a\",1/0)")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=SUBSTITUTE(1/0,\"a\",\"b\")")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(Calc("=CHAR(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }
}
