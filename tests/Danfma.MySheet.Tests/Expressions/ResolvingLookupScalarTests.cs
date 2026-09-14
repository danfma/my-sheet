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
        main["Z1"] = ExpressionParser.Parse("=1/0", main);
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
}
