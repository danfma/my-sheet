using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

public class BlankLookupKeyTests
{
    [Test]
    [Arguments("=MATCH(D1,B1:B4,0)", "1")]
    [Arguments("=MATCH(D1,B1:B4,1)", "1")]
    [Arguments("=MATCH(D1,B1:B4,-1)", "1")]
    [Arguments("=XMATCH(D1,B1:B4)", "3")]
    [Arguments("=XMATCH(D1,B1:B4,0,-1)", "2")]
    [Arguments("=XLOOKUP(D1,B1:B4,C1:C4)", "30")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,FALSE)", "10")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,TRUE)", "10")]
    [Arguments("=COUNTIF(B1:B4,D1)", "1")]
    [Arguments("=COUNTIFS(B1:B4,D1)", "1")]
    [Arguments("=SUMIF(B1:B4,D1,C1:C4)", "10")]
    [Arguments("=AVERAGEIF(B1:B4,D1,C1:C4)", "10")]
    [Arguments("=MAXIFS(C1:C4,B1:B4,D1)", "10")]
    [Arguments("=LET(k,D1,XMATCH(k,B1:B4))", "3")]
    public async Task AbsentKey_UsesTheConsumerFamilyRule(string formula, string expected) =>
        await Assert.That(Evaluate(formula, Key.Absent, mixed: true)).IsEqualTo(expected);

    [Test]
    [Arguments("=XLOOKUP(D1,B1:B4,C1:C4)", "#N/A")]
    [Arguments("=MATCH(D1,B1:B4,0)", "#N/A")]
    [Arguments("=COUNTIF(B1:B4,D1)", "0")]
    public async Task AbsentKey_WithNoBlankOrZero_HasNoMatch(string formula, string expected) =>
        await Assert.That(Evaluate(formula, Key.Absent, mixed: false)).IsEqualTo(expected);

    [Test]
    [Arguments("=MATCH(D1,B1:B4,0)", "2")]
    [Arguments("=MATCH(D1,B1:B4,1)", "2")]
    [Arguments("=MATCH(D1,B1:B4,-1)", "2")]
    [Arguments("=XMATCH(D1,B1:B4)", "2")]
    [Arguments("=XLOOKUP(D1,B1:B4,C1:C4)", "20")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,FALSE)", "20")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,TRUE)", "20")]
    [Arguments("=COUNTIF(B1:B4,D1)", "2")]
    [Arguments("=SUMIF(B1:B4,D1,C1:C4)", "50")]
    [Arguments("=COUNTIFS(B1:B4,D1)", "2")]
    public async Task EmptyTextKey_RemainsDistinctFromAbsent(string formula, string expected) =>
        await Assert.That(Evaluate(formula, Key.FormulaEmpty, mixed: true)).IsEqualTo(expected);

    [Test]
    [Arguments("=MATCH(D1,B1:B4,1)", "#N/A")]
    [Arguments("=MATCH(D1,B1:B4,-1)", "#N/A")]
    public async Task ApproximateMatch_EmptyTextWithoutMatch_IsNotAvailable(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, Key.ExplicitEmpty, mixed: false)).IsEqualTo(expected);

    [Test]
    [Arguments("=COUNTIF(B1:B4,\"\")", "2")]
    [Arguments("=COUNTIF(B1:B4,\"=\")", "1")]
    [Arguments("=COUNTIF(B1:B4,\"<>\")", "3")]
    public async Task EmptyCriteria_DistinguishImplicitAndExplicitEquality(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, Key.Absent, mixed: true)).IsEqualTo(expected);

    [Test]
    [Arguments("=HLOOKUP(D1,B1:E2,2,FALSE)", "10", Key.Absent)]
    [Arguments("=HLOOKUP(D1,B1:E2,2,TRUE)", "10", Key.Absent)]
    [Arguments("=HLOOKUP(D1,B1:E2,2,FALSE)", "20", Key.FormulaEmpty)]
    [Arguments("=HLOOKUP(D1,B1:E2,2,TRUE)", "30", Key.FormulaEmpty)]
    [Arguments("=HLOOKUP(D1,B1:E2,2,FALSE)", "20", Key.ExplicitEmpty)]
    [Arguments("=HLOOKUP(D1,B1:E2,2,TRUE)", "30", Key.ExplicitEmpty)]
    public async Task HLookup_BlankKeys_MatchTheTransposedFixture(
        string formula,
        string expected,
        Key key
    ) => await Assert.That(EvaluateHorizontal(formula, key)).IsEqualTo(expected);

    [Test]
    [Arguments("=VLOOKUP(D1,{0,10;\"\",20;\"x\",30;5,40},2,FALSE)", "10")]
    [Arguments("=VLOOKUP(D1,{0,10;\"\",20;\"x\",30;5,40},2,TRUE)", "10")]
    [Arguments("=HLOOKUP(D1,{0,\"\",\"x\",5;10,20,30,40},2,FALSE)", "10")]
    [Arguments("=HLOOKUP(D1,{0,\"\",\"x\",5;10,20,30,40},2,TRUE)", "10")]
    public async Task TableLookups_ArrayRoute_PreservesAbsentZeroRule(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, Key.Absent, mixed: true)).IsEqualTo(expected);

    [Test]
    public async Task AbsentKey_OnAnotherSheet_UsesTheSameRule()
    {
        var workbook = CreateWorkbook(Key.Absent, mixed: true);
        workbook.Sheets.Add("Other");
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse("=XMATCH(Other!D1,B1:B4)", main);

        await Assert.That(Format(workbook.GetCellValue("Main", "AZ5000"))).IsEqualTo("3");
    }

    private static string Evaluate(string formula, Key key, bool mixed)
    {
        var workbook = CreateWorkbook(key, mixed);
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return Format(workbook.GetCellValue("Main", "AZ5000"));
    }

    private static Workbook CreateWorkbook(Key key, bool mixed)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["B1"] = new NumberValue(mixed ? 0 : 1);
        main["B2"] = mixed ? ExpressionParser.Parse("=\"\"", main) : new NumberValue(2);
        if (!mixed)
        {
            main["B3"] = new NumberValue(3);
        }
        main["B4"] = new NumberValue(5);
        for (var row = 1; row <= 4; row++)
        {
            main[$"C{row}"] = new NumberValue(row * 10);
        }
        SetKey(main, key);
        return workbook;
    }

    private static string EvaluateHorizontal(string formula, Key key)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["B1"] = new NumberValue(0);
        main["C1"] = ExpressionParser.Parse("=\"\"", main);
        main["E1"] = new NumberValue(5);
        main["B2"] = new NumberValue(10);
        main["C2"] = new NumberValue(20);
        main["D2"] = new NumberValue(30);
        main["E2"] = new NumberValue(40);
        SetKey(main, key);
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return Format(workbook.GetCellValue("Main", "AZ5000"));
    }

    private static void SetKey(Sheet sheet, Key key)
    {
        if (key == Key.FormulaEmpty)
        {
            sheet["D1"] = ExpressionParser.Parse("=\"\"", sheet);
        }
        else if (key == Key.ExplicitEmpty)
        {
            sheet["D1"] = new Danfma.MySheet.Expressions.StringValue(string.Empty);
        }
    }

    private static string Format(ComputedValue value) =>
        value.TryGetError(out var error) ? error.ToString()
        : value.TryGetNumber(out var number) ? number.ToString(CultureInfo.InvariantCulture)
        : value.TryGetText(out var text) ? $"\"{text}\""
        : value.Kind.ToString();

    public enum Key
    {
        Absent,
        FormulaEmpty,
        ExplicitEmpty,
    }
}
