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

        var expression = ExpressionParser.Parse(formula, main);
        var value = expression.Evaluate(new EvaluationContext(workbook, "Main", "AZ5000"));

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

    // Fixture: A1:A3 = 5,0,9; B1:B3 = 1,2,3; C1:C3 = 10,20,30; formula at AZ5000.
    // Aspose.Cells 26.7.0 PLAIN / CSE agree: mismatched axes and 2D lookup arrays are #VALUE!,
    // while equal computed vertical and horizontal vectors execute and return their first element.
    // Before this fix every computed row below returned 0 and each 2D row returned 10 in MySheet.
    [Test]
    [Arguments("=XLOOKUP(1,SEQUENCE(3),B1:B2)", "#VALUE!")]
    [Arguments("=XLOOKUP(5,A1:A3,SEQUENCE(2))", "#VALUE!")]
    [Arguments("=XLOOKUP(1,SEQUENCE(3),SEQUENCE(2))", "#VALUE!")]
    [Arguments("=XLOOKUP(1,SEQUENCE(3),SEQUENCE(3))", "1")]
    [Arguments("=XLOOKUP(5,A1:B2,C1:C2)", "#VALUE!")]
    [Arguments("=XLOOKUP(5,A1:B2,C1:D2)", "#VALUE!")]
    [Arguments("=XLOOKUP(1,SEQUENCE(1,3),B1:C1)", "#VALUE!")]
    [Arguments("=XLOOKUP(5,A1:C1,SEQUENCE(1,2))", "#VALUE!")]
    [Arguments("=XLOOKUP(1,SEQUENCE(1,3),SEQUENCE(1,2))", "#VALUE!")]
    [Arguments("=XLOOKUP(1,SEQUENCE(1,3),SEQUENCE(1,3))", "1")]
    public async Task XLookup_ValidatesAndExecutesReferenceOrComputedShapes(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    [Test]
    public async Task XLookup_ComputedOperands_AreBuiltExactlyOnce()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        var draws = 0;
        workbook.RegisterFunction("TICK", (_, _) => ++draws);

        var value = ExpressionParser
            .Parse("=XLOOKUP(1,SEQUENCE(3,1,TICK(),0),SEQUENCE(3))", sheet)
            .Evaluate(new EvaluationContext(workbook, "Main", "AZ5000"));

        await Assert.That(draws).IsEqualTo(1);
        await Assert.That(value).IsEqualTo(ComputedValue.Number(1));
    }

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
