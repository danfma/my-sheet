using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

/// <summary>
/// FIXED/DOLLAR/NUMBERVALUE/TEXTBEFORE/TEXTAFTER/VALUETOTEXT. Oracles: the official Microsoft
/// support pages for each function (fetched 2026-07-01), cited per test. FIXED/DOLLAR follow the
/// engine's locale-invariant contract (§A7): '.' decimal, ',' thousands, '$' symbol.
/// </summary>
public class TextFormattingTests
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
    public async Task Fixed_MatchesExcelDocs()
    {
        // support.microsoft.com FIXED (A2=1234.567, A3=-1234.567, A4=44.332):
        // FIXED(A2,1)="1,234.6"; FIXED(A2,-1)="1,230"; FIXED(A3,-1,TRUE)="-1230";
        // FIXED(A4)="44.33" (decimals defaults to 2).
        await Assert.That(Calc("=FIXED(1234.567,1)") as string).IsEqualTo("1,234.6");
        await Assert.That(Calc("=FIXED(1234.567,-1)") as string).IsEqualTo("1,230");
        await Assert.That(Calc("=FIXED(-1234.567,-1,TRUE)") as string).IsEqualTo("-1230");
        await Assert.That(Calc("=FIXED(44.332)") as string).IsEqualTo("44.33");
    }

    [Test]
    public async Task Fixed_IsTextNotANumber()
    {
        await Assert.That(Calc("=ISTEXT(FIXED(1))") as bool?).IsTrue();
        // The doc caps decimals at 127.
        await Assert.That(Calc("=FIXED(1,200)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task Dollar_MatchesExcelDocs()
    {
        // support.microsoft.com DOLLAR (A2=1234.567, A3=-1234.567, A4=-0.123, A5=99.888):
        // DOLLAR(A2,2)="$1,234.57"; DOLLAR(A3,-2)="($1,200)"; DOLLAR(A4,4)="($0.1230)";
        // DOLLAR(A5)="$99.89"; DOLLAR(A2,-2)="$1,200" (older text table of the same page).
        // Negatives use parentheses — the documented $#,##0.00_);($#,##0.00) format.
        await Assert.That(Calc("=DOLLAR(1234.567,2)") as string).IsEqualTo("$1,234.57");
        await Assert.That(Calc("=DOLLAR(-1234.567,-2)") as string).IsEqualTo("($1,200)");
        await Assert.That(Calc("=DOLLAR(-0.123,4)") as string).IsEqualTo("($0.1230)");
        await Assert.That(Calc("=DOLLAR(99.888)") as string).IsEqualTo("$99.89");
        await Assert.That(Calc("=DOLLAR(1234.567,-2)") as string).IsEqualTo("$1,200");
    }

    [Test]
    public async Task NumberValue_MatchesExcelDocs()
    {
        // support.microsoft.com NUMBERVALUE: NUMBERVALUE("2.500,27",",",".")=2500.27;
        // NUMBERVALUE("3.5%")=0.035; "9%%"=0.0009 (percent signs compound); spaces are ignored
        // even in the middle (" 3 000 " -> 3000); empty text -> 0.
        await Assert
            .That(Calc("=NUMBERVALUE(\"2.500,27\",\",\",\".\")") as double?)
            .IsEqualTo(2500.27);
        await Assert.That(Calc("=NUMBERVALUE(\"3.5%\")") as double?).IsEqualTo(0.035);
        await Assert.That(Calc("=NUMBERVALUE(\"9%%\")") as double?).IsEqualTo(0.0009);
        await Assert.That(Calc("=NUMBERVALUE(\" 3 000 \")") as double?).IsEqualTo(3000.0);
        await Assert.That(Calc("=NUMBERVALUE(\"\")") as double?).IsEqualTo(0.0);
    }

    [Test]
    public async Task NumberValue_ErrorCases()
    {
        // support.microsoft.com NUMBERVALUE remarks: a decimal separator used more than once ->
        // #VALUE!; a group separator occurring after the decimal separator -> #VALUE! (before it,
        // the group separator is simply ignored); invalid text -> #VALUE!.
        await Assert.That(Calc("=NUMBERVALUE(\"1.2.3\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=NUMBERVALUE(\"3.5,0\")")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=NUMBERVALUE(\"1,234.5\")") as double?).IsEqualTo(1234.5);
        await Assert.That(Calc("=NUMBERVALUE(\"abc\")")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task TextBefore_MatchesExcelDocs()
    {
        // support.microsoft.com TEXTBEFORE (A2="Little Red Riding Hood's red hood",
        // A3="Little red Riding Hood's red hood"): TEXTBEFORE(A2,"Red")="Little ";
        // TEXTBEFORE(A3,"Red")=#N/A (case-sensitive by default); TEXTBEFORE(A3,"red",2)=
        // "Little red Riding Hood's "; TEXTBEFORE(A3,"red",-2)="Little ";
        // TEXTBEFORE(A3,"red",3)=#N/A.
        var a2 = ("A2", Expression.String("Little Red Riding Hood's red hood"));
        var a3 = ("A3", Expression.String("Little red Riding Hood's red hood"));

        await Assert.That(Calc("=TEXTBEFORE(A2,\"Red\")", a2) as string).IsEqualTo("Little ");
        await Assert.That(Calc("=TEXTBEFORE(A3,\"Red\")", a3)).IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Calc("=TEXTBEFORE(A3,\"red\",2)", a3) as string)
            .IsEqualTo("Little red Riding Hood's ");
        await Assert.That(Calc("=TEXTBEFORE(A3,\"red\",-2)", a3) as string).IsEqualTo("Little ");
        await Assert.That(Calc("=TEXTBEFORE(A3,\"red\",3)", a3)).IsEqualTo(ErrorValue.NotAvailable);
        // Case-insensitive match_mode 1 finds the first "red" regardless of case.
        await Assert.That(Calc("=TEXTBEFORE(A3,\"Red\",1,1)", a3) as string).IsEqualTo("Little ");
    }

    [Test]
    public async Task TextBefore_MatchEndAndDefaults()
    {
        // support.microsoft.com TEXTBEFORE example 2: TEXTBEFORE("Marcus Aurelius"," ",,,1)=
        // "Marcus"; TEXTBEFORE("Socrates"," ",,,0)=#N/A; TEXTBEFORE("Socrates"," ",,,1)=
        // "Socrates" (the end of text counts as the delimiter).
        await Assert
            .That(Calc("=TEXTBEFORE(\"Marcus Aurelius\",\" \",,,1)") as string)
            .IsEqualTo("Marcus");
        await Assert
            .That(Calc("=TEXTBEFORE(\"Socrates\",\" \",,,0)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert
            .That(Calc("=TEXTBEFORE(\"Socrates\",\" \",,,1)") as string)
            .IsEqualTo("Socrates");

        // Empty delimiter matches immediately: "" from the front, the whole text from the end.
        await Assert
            .That(Calc("=TEXTBEFORE(\"Red riding hood's, red hood\",\"\")") as string)
            .IsEqualTo("");
        await Assert
            .That(Calc("=TEXTBEFORE(\"Red riding hood's, red hood\",\"\",-1)") as string)
            .IsEqualTo("Red riding hood's, red hood");

        // instance_num = 0 or beyond the text length -> #VALUE!; if_not_found overrides #N/A.
        await Assert.That(Calc("=TEXTBEFORE(\"abc\",\"b\",0)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=TEXTBEFORE(\"abc\",\"b\",99)")).IsEqualTo(ErrorValue.NotValue);
        await Assert
            .That(Calc("=TEXTBEFORE(\"abc\",\"x\",1,0,0,\"none\")") as string)
            .IsEqualTo("none");
    }

    [Test]
    public async Task TextAfter_MatchesExcelDocs()
    {
        // support.microsoft.com TEXTAFTER (same A2/A3 data): TEXTAFTER(A2,"Red")=
        // " Riding Hood's red hood"; TEXTAFTER(A2,"basket")=#N/A; TEXTAFTER(A3,"red",2)=" hood";
        // TEXTAFTER(A3,"red",-2)=" Riding Hood's red hood"; TEXTAFTER(A2,"red",3)=#N/A.
        var a2 = ("A2", Expression.String("Little Red Riding Hood's red hood"));
        var a3 = ("A3", Expression.String("Little red Riding Hood's red hood"));

        await Assert
            .That(Calc("=TEXTAFTER(A2,\"Red\")", a2) as string)
            .IsEqualTo(" Riding Hood's red hood");
        await Assert.That(Calc("=TEXTAFTER(A2,\"basket\")", a2)).IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=TEXTAFTER(A3,\"red\",2)", a3) as string).IsEqualTo(" hood");
        await Assert
            .That(Calc("=TEXTAFTER(A3,\"red\",-2)", a3) as string)
            .IsEqualTo(" Riding Hood's red hood");
        await Assert.That(Calc("=TEXTAFTER(A2,\"red\",3)", a2)).IsEqualTo(ErrorValue.NotAvailable);
    }

    [Test]
    public async Task TextAfter_MatchEndAndDefaults()
    {
        // support.microsoft.com TEXTAFTER example 2: TEXTAFTER("Marcus Aurelius"," ",,,1)=
        // "Aurelius"; TEXTAFTER("Socrates"," ",,,0)=#N/A; TEXTAFTER("Socrates"," ",,,1)="" (the
        // end of text counts as the delimiter). Empty delimiter: whole text from the front, ""
        // from the end.
        await Assert
            .That(Calc("=TEXTAFTER(\"Marcus Aurelius\",\" \",,,1)") as string)
            .IsEqualTo("Aurelius");
        await Assert
            .That(Calc("=TEXTAFTER(\"Socrates\",\" \",,,0)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=TEXTAFTER(\"Socrates\",\" \",,,1)") as string).IsEqualTo("");
        await Assert
            .That(Calc("=TEXTAFTER(\"Red riding hood's, red hood\",\"\")") as string)
            .IsEqualTo("Red riding hood's, red hood");
        await Assert
            .That(Calc("=TEXTAFTER(\"Red riding hood's, red hood\",\"\",-1)") as string)
            .IsEqualTo("");
        await Assert
            .That(Calc("=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\")") as string)
            .IsEqualTo("'s, red hood");
    }

    [Test]
    public async Task ValueToText_MatchesExcelDocs()
    {
        // support.microsoft.com VALUETOTEXT: concise (0, the default) renders like the cell —
        // TRUE, 1234.01234, Hello, #VALUE!; strict (1) wraps text in quotes but not Booleans,
        // Numbers or Errors. Any other format -> #VALUE!.
        await Assert.That(Calc("=VALUETOTEXT(TRUE,0)") as string).IsEqualTo("TRUE");
        await Assert.That(Calc("=VALUETOTEXT(1234.01234,0)") as string).IsEqualTo("1234.01234");
        await Assert.That(Calc("=VALUETOTEXT(\"Hello\",0)") as string).IsEqualTo("Hello");
        await Assert.That(Calc("=VALUETOTEXT(TRUE,1)") as string).IsEqualTo("TRUE");
        await Assert.That(Calc("=VALUETOTEXT(1234,1)") as string).IsEqualTo("1234");
        await Assert.That(Calc("=VALUETOTEXT(\"Hello\",1)") as string).IsEqualTo("\"Hello\"");
        await Assert.That(Calc("=VALUETOTEXT(NA(),0)") as string).IsEqualTo("#N/A");
        await Assert.That(Calc("=VALUETOTEXT(NA(),1)") as string).IsEqualTo("#N/A");
        await Assert.That(Calc("=VALUETOTEXT(\"x\",2)")).IsEqualTo(ErrorValue.NotValue);
        await Assert.That(Calc("=VALUETOTEXT(\"x\")") as string).IsEqualTo("x"); // default 0
    }
}
