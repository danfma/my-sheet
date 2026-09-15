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

    // Aspose 26.7.0 PLAIN/CSE resolves each key to the absent D1 before classifying it. Before this fix,
    // XMATCH was 1 -> 3, approximate MATCH 4 -> 1, XLOOKUP "" -> 30, COUNTIF 2 -> 1, and SUMIF 50 -> 0.
    [Test]
    [Arguments("=XMATCH(INDEX(D1:D1,1),B1:B4)", "3")]
    [Arguments("=XMATCH(OFFSET(D1,0,0),B1:B4)", "3")]
    [Arguments("=XMATCH(IF(TRUE,D1),B1:B4)", "3")]
    [Arguments("=XMATCH(EmptyCell,B1:B4)", "3")]
    [Arguments("=MATCH(INDEX(D1:D1,1),B1:B4,-1)", "1")]
    [Arguments("=MATCH(OFFSET(D1,0,0),B1:B4,-1)", "1")]
    [Arguments("=MATCH(IF(TRUE,D1),B1:B4,-1)", "1")]
    [Arguments("=MATCH(EmptyCell,B1:B4,-1)", "1")]
    [Arguments("=XLOOKUP(INDEX(D1:D1,1),B1:B4,C1:C4)", "30")]
    [Arguments("=XLOOKUP(OFFSET(D1,0,0),B1:B4,C1:C4)", "30")]
    [Arguments("=XLOOKUP(IF(TRUE,D1),B1:B4,C1:C4)", "30")]
    [Arguments("=XLOOKUP(EmptyCell,B1:B4,C1:C4)", "30")]
    [Arguments("=COUNTIF(B1:B4,INDEX(D1:D1,1))", "1")]
    [Arguments("=COUNTIF(B1:B4,OFFSET(D1,0,0))", "1")]
    [Arguments("=COUNTIF(B1:B4,IF(TRUE,D1))", "1")]
    [Arguments("=COUNTIF(B1:B4,EmptyCell)", "1")]
    [Arguments("=SUMIF(B1:B4,INDEX(D1:D1,1),C1:C4)", "0")]
    [Arguments("=SUMIF(B1:B4,OFFSET(D1,0,0),C1:C4)", "0")]
    [Arguments("=SUMIF(B1:B4,IF(TRUE,D1),C1:C4)", "0")]
    [Arguments("=SUMIF(B1:B4,EmptyCell,C1:C4)", "0")]
    public async Task DerivedAbsentKey_UsesTheConsumerFamilyRule(string formula, string expected) =>
        await Assert.That(Evaluate(formula, Key.Absent, mixed: true)).IsEqualTo(expected);

    [Test]
    public async Task IndexSelectingAnAbsentCellInsideARange_IsAnAbsentKey() =>
        // Before this fix the empty-text-equivalent path returned 1; the oracle returns the blank at 3.
        await Assert
            .That(Evaluate("=XMATCH(INDEX(B1:D1,1,3),B1:B4)", Key.Absent, mixed: true))
            .IsEqualTo("3");

    [Test]
    [Arguments(false, false, "#N/A")]
    [Arguments(true, false, "#N/A")]
    [Arguments(false, true, "30")]
    [Arguments(true, true, "30")]
    public async Task ApproximateTableLookup_EmptyTextRequiresAndUsesTheLastExactCandidate(
        bool horizontal,
        bool severalEmptyTexts,
        string expected
    ) =>
        // No candidate was 30 -> #N/A; changing last-match to first makes the multi-text 30 -> 20.
        await Assert
            .That(EvaluateApproximateTableLookup(horizontal, severalEmptyTexts))
            .IsEqualTo(expected);

    [Test]
    [Arguments("=XMATCH(OFFSET(D1,TICK()*0,0),B1:B4)", "3")]
    [Arguments("=COUNTIF(B1:B4,OFFSET(D1,TICK()*0,0))", "1")]
    public async Task DerivedAbsentKey_IsResolvedOnlyOnce(string formula, string expected)
    {
        var draws = 0;
        var workbook = CreateWorkbook(Key.Absent, mixed: true);
        workbook.RegisterFunction("TICK", (_, _) => ++draws);
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse(formula, main);

        await Assert.That(Format(workbook.GetCellValue("Main", "AZ5000"))).IsEqualTo(expected);
        await Assert.That(draws).IsEqualTo(1);
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
        workbook.DefineName("EmptyCell", "Main!D1");
        SetKey(main, key);
        return workbook;
    }

    private static string EvaluateApproximateTableLookup(bool horizontal, bool severalEmptyTexts)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        var keys = severalEmptyTexts
            ? new Expression?[]
            {
                new NumberValue(0),
                ExpressionParser.Parse("=\"\"", main),
                ExpressionParser.Parse("=\"\"", main),
                null,
                new NumberValue(5),
            }
            : new Expression?[] { new NumberValue(0), null, new NumberValue(5) };

        for (var i = 0; i < keys.Length; i++)
        {
            var position = i + 1;
            var keyCell = horizontal ? $"{(char)('A' + position)}1" : $"B{position}";
            var valueCell = horizontal ? $"{(char)('A' + position)}2" : $"C{position}";
            if (keys[i] is { } key)
            {
                main[keyCell] = key;
            }
            main[valueCell] = new NumberValue(position * 10);
        }

        main["H10"] = new Danfma.MySheet.Expressions.StringValue(string.Empty);
        var formula = horizontal ? "=HLOOKUP(H10,B1:F2,2,TRUE)" : "=VLOOKUP(H10,B1:C5,2,TRUE)";
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return Format(workbook.GetCellValue("Main", "AZ5000"));
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
