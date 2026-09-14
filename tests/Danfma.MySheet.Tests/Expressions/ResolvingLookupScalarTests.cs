using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class ResolvingLookupScalarTests
{
    private static string Evaluate(string formula)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(1);
        main["Z1"] = ExpressionParser.Parse("=1/0", main);
        workbook.DefineName("ValueCell", "Main!$A$1");
        workbook.DefineName("ErrCell", "Main!$Z$1");
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

        return value.Kind.ToString();
    }

    [Test]
    [Arguments("=VLOOKUP(1,5,1)", "#N/A")]
    [Arguments("=VLOOKUP(1,5,2)", "#REF!")]
    [Arguments("=VLOOKUP(1,ErrCell,1)", "#DIV/0!")]
    [Arguments("=VLOOKUP(1,1/0,1)", "#N/A")]
    [Arguments("=HLOOKUP(1,5,1)", "#N/A")]
    [Arguments("=HLOOKUP(1,5,2)", "#REF!")]
    [Arguments("=HLOOKUP(1,ErrCell,1)", "#DIV/0!")]
    [Arguments("=HLOOKUP(1,1/0,1)", "#N/A")]
    [Arguments("=INDEX(5,1)", "5")]
    [Arguments("=INDEX(ErrCell,1)", "#DIV/0!")]
    [Arguments("=INDEX(1/0,1)", "#DIV/0!")]
    [Arguments("=MATCH(1,5,0)", "#N/A")]
    [Arguments("=MATCH(1,ErrCell,0)", "#N/A")]
    [Arguments("=MATCH(1,1/0,0)", "#N/A")]
    public async Task ResolvingLookups_ApplyTheirScalarFallback(string formula, string expected) =>
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    // Fixture: Main!A1 = 1, ValueCell = Main!$A$1, formula at Main!AZ5000. Aspose.Cells
    // 26.7.0 PLAIN / CSE agree on every row: successful 1x1 lookups return 1; index 0 is #VALUE!,
    // an index beyond the 1x1 table is #REF!, and an invalid mode propagates #DIV/0!.
    [Test]
    [Arguments("=VLOOKUP(1,A1,1,FALSE)", "1")]
    [Arguments("=VLOOKUP(1,A1,2,FALSE)", "#REF!")]
    [Arguments("=VLOOKUP(1,A1,0,FALSE)", "#VALUE!")]
    [Arguments("=VLOOKUP(1,A1,1,1/0)", "#DIV/0!")]
    [Arguments("=VLOOKUP(1,ValueCell,1,FALSE)", "1")]
    [Arguments("=HLOOKUP(1,A1,1,FALSE)", "1")]
    [Arguments("=HLOOKUP(1,A1,2,FALSE)", "#REF!")]
    [Arguments("=HLOOKUP(1,A1,0,FALSE)", "#VALUE!")]
    [Arguments("=HLOOKUP(1,A1,1,1/0)", "#DIV/0!")]
    [Arguments("=HLOOKUP(1,ValueCell,1,FALSE)", "1")]
    [Arguments("=INDEX(A1,2)", "#REF!")]
    [Arguments("=INDEX(A1,1,2)", "#REF!")]
    [Arguments("=INDEX(A1,1,1)", "1")]
    [Arguments("=INDEX(A1,0)", "1")]
    [Arguments("=INDEX(A1,0,1)", "1")]
    [Arguments("=INDEX(ValueCell,1)", "1")]
    public async Task CellReferences_AreValidatedAsOneByOneTables(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula)).IsEqualTo(expected);
}
