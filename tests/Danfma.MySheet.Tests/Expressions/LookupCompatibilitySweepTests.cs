using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

public class LookupCompatibilitySweepTests
{
    private static string Evaluate(string formula, bool withTable = false)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(5);
        main["A2"] = new NumberValue(0);
        main["A3"] = new NumberValue(9);
        main["B1"] = new NumberValue(1);
        main["B2"] = new NumberValue(2);
        main["C1"] = new NumberValue(10);
        main["C2"] = new NumberValue(20);
        main["C3"] = new NumberValue(30);
        main["E1"] = new NumberValue(5);
        main["F1"] = new NumberValue(0);
        main["G1"] = new NumberValue(9);
        main["E2"] = new NumberValue(1);
        main["F2"] = new NumberValue(2);
        if (withTable)
        {
            var data = workbook.Sheets.Add("Data");
            data["A1"] = new StringValue("Item");
            data["B1"] = new StringValue("Valor");
            data["C1"] = new StringValue("Qtd");
            data["A2"] = new StringValue("a");
            data["B2"] = new NumberValue(10);
            data["C2"] = new NumberValue(1);
            data["A3"] = new StringValue("b");
            data["B3"] = new NumberValue(20);
            data["C3"] = new NumberValue(2);
            data["A4"] = new StringValue("c");
            data["B4"] = new NumberValue(30);
            data["C4"] = new NumberValue(3);
            workbook.DefineTable("Tabela1", "Data", "A1:C4", ["Item", "Valor", "Qtd"]);
        }

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

        return value.TryGetText(out var text) ? $"\"{text}\"" : value.Kind.ToString();
    }

    [Test]
    [Arguments("=XLOOKUP(5,A1:A3,B1:B2)", "#VALUE!")]
    [Arguments("=XLOOKUP(9,A1:A3,B1:B2)", "#VALUE!")]
    [Arguments("=XLOOKUP(5,A1:A3,B1:C3)", "1")]
    [Arguments("=XLOOKUP(5,E1:G1,E2:F2)", "#VALUE!")]
    [Arguments("=XLOOKUP(9,E1:G1,E2:F2)", "#VALUE!")]
    [Arguments("=XLOOKUP(5,A1:A3,B1:B2,\"nf\")", "#VALUE!")]
    [Arguments("=XLOOKUP(9,A1:A3,B1:B2,\"nf\")", "#VALUE!")]
    public async Task XLookup_RequiresMatchingLookupAxis(string formula, string expected) =>
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    [Test]
    [Arguments("=XLOOKUP(1,NoSuch,B1:B3)", "#N/A")]
    [Arguments("=XLOOKUP(1,A1:A3,NoSuch)", "#VALUE!")]
    [Arguments("=XLOOKUP(1,Tabela1[#Totals],B1:B3)", "#REF!")]
    [Arguments("=XLOOKUP(1,A1:A3,Tabela1[#Totals])", "#REF!")]
    public async Task XLookup_UnresolvableArraySlots_UseSlotFallbackOrStructuralError(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, withTable: true)).IsEqualTo(expected);
}
