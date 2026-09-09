using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class TextFormatTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    public async Task Text_NumericFormats()
    {
        await Assert.That(Calc("=TEXT(1234.5,\"#,##0.00\")") as string).IsEqualTo("1,234.50");
        await Assert.That(Calc("=TEXT(0.5,\"0%\")") as string).IsEqualTo("50%");
    }

    [Test]
    public async Task Text_DateFormats()
    {
        // 44197 is the Excel serial for 2021-01-01.
        await Assert.That(Calc("=TEXT(44197,\"yyyy-mm-dd\")") as string).IsEqualTo("2021-01-01");
        await Assert.That(Calc("=TEXT(44197,\"dd/mm/yyyy\")") as string).IsEqualTo("01/01/2021");
        await Assert.That(Calc("=TEXT(44197,\"mmm yyyy\")") as string).IsEqualTo("Jan 2021");
    }

    [Test]
    public async Task Text_TimeFormats()
    {
        // 0.5 of a day is noon; mm here is minutes (it follows hh).
        await Assert.That(Calc("=TEXT(0.5,\"hh:mm:ss\")") as string).IsEqualTo("12:00:00");
    }

    // The FIELD matrix, on a modern serial where no epoch question is in play: 45366 = Friday 2024-03-15.
    // MEASURED on Aspose.Cells 26.6.0 on 2026-09-09, PLAIN cell entry (a formula assigned to a cell), and
    // every row below is byte-identical to it. The 1900-window rows of the same matrix live in DateEpochTests,
    // which owns the epoch; this test owns the format CODES.
    [Test]
    public async Task Text_DateFieldMatrix_OnAModernSerial()
    {
        await Assert.That(Calc("=TEXT(45366,\"yyyy\")") as string).IsEqualTo("2024");
        await Assert.That(Calc("=TEXT(45366,\"yy\")") as string).IsEqualTo("24");
        await Assert.That(Calc("=TEXT(45366,\"mm\")") as string).IsEqualTo("03");
        await Assert.That(Calc("=TEXT(45366,\"mmm\")") as string).IsEqualTo("Mar");
        await Assert.That(Calc("=TEXT(45366,\"mmmm\")") as string).IsEqualTo("March");
        await Assert.That(Calc("=TEXT(45366,\"dd\")") as string).IsEqualTo("15");
        await Assert.That(Calc("=TEXT(45366,\"ddd\")") as string).IsEqualTo("Fri");
        await Assert.That(Calc("=TEXT(45366,\"dddd\")") as string).IsEqualTo("Friday");
        await Assert
            .That(Calc("=TEXT(45366,\"dddd d mmmm yyyy\")") as string)
            .IsEqualTo("Friday 15 March 2024");
        await Assert.That(Calc("=TEXT(45366,\"d-mmm\")") as string).IsEqualTo("15-Mar");
        await Assert
            .That(Calc("=TEXT(45366,\"mmmm d, yyyy\")") as string)
            .IsEqualTo("March 15, 2024");
    }

    [Test]
    public async Task Text_TimeFieldsRideTheFraction()
    {
        // MEASURED on Aspose.Cells 26.6.0 on 2026-09-09, PLAIN cell entry. The fraction is read as a whole
        // second count, so a date field and a time field in one format agree on the same instant.
        await Assert
            .That(Calc("=TEXT(45366.5,\"yyyy-mm-dd hh:mm\")") as string)
            .IsEqualTo("2024-03-15 12:00");
        await Assert.That(Calc("=TEXT(0.75,\"hh:mm:ss\")") as string).IsEqualTo("18:00:00");
        await Assert.That(Calc("=TEXT(0,\"hh:mm:ss\")") as string).IsEqualTo("00:00:00");
        await Assert.That(Calc("=TEXT(0.5,\"h:mm AM/PM\")") as string).IsEqualTo("12:00 PM");
        // "mm" with no hour before it and no seconds after it is a MONTH, even on a fractional serial.
        await Assert.That(Calc("=TEXT(45366.5,\"mm\")") as string).IsEqualTo("03");
        await Assert.That(Calc("=TEXT(45366.5,\"hh:mm mmmm\")") as string).IsEqualTo("12:00 March");
    }

    // A format that is ONE letter is an OPEN QUESTION in Phase 9, recorded in the phase file as a P0 work
    // item: Aspose answers the plain field (TEXT(45366,"d") = 15, "m" = 3, "y" = 24, "s" = 0, "h" = 0) and so
    // does MySheet now, but the two still disagree inside the 1900 window, where Aspose reads a lone d/m/y
    // off the DateTime map with no day-zero and no phantom rule: TEXT(0,"d") = 31 and TEXT(60,"d") = 28
    // against 0 and 29 here. Pinned as MySheet answers so the open question cannot move silently.
    [Test]
    public async Task Text_LoneFieldTokens()
    {
        await Assert.That(Calc("=TEXT(45366,\"d\")") as string).IsEqualTo("15");
        await Assert.That(Calc("=TEXT(45366,\"m\")") as string).IsEqualTo("3");
        await Assert.That(Calc("=TEXT(45366,\"y\")") as string).IsEqualTo("24");
        await Assert.That(Calc("=TEXT(45366,\"d/m\")") as string).IsEqualTo("15/3");
        await Assert.That(Calc("=TEXT(45366,\"m/d\")") as string).IsEqualTo("3/15");
        await Assert.That(Calc("=TEXT(45366.5,\"h\")") as string).IsEqualTo("12");
        await Assert.That(Calc("=TEXT(45366.5,\"s\")") as string).IsEqualTo("0");
    }

    [Test]
    public async Task Text_LiteralsAndEscapes()
    {
        // Quoted runs and backslash escapes print verbatim, so a letter inside them is never a field — and a
        // quoted "dddd" must not trip the weekday-name exception that decides the phantom Feb 29 either.
        await Assert.That(Calc("=TEXT(45366,\"dd\\d\")") as string).IsEqualTo("15d");
        await Assert
            .That(Calc("=TEXT(45366,\"dd \"\"dddd\"\" mm\")") as string)
            .IsEqualTo("15 dddd 03");
        await Assert
            .That(Calc("=TEXT(60,\"yyyy-mm-dd \"\"dddd\"\"\")") as string)
            .IsEqualTo("1900-02-29 dddd");
    }
}
