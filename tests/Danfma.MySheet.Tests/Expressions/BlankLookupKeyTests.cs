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
    [Arguments("=XLOOKUP(D1,B1:B4,C1:C4)", "4")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,FALSE)", "1")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,TRUE)", "1")]
    [Arguments("=COUNTIF(B1:B4,D1)", "1")]
    [Arguments("=COUNTIFS(B1:B4,D1)", "1")]
    [Arguments("=SUMIF(B1:B4,D1,C1:C4)", "1")]
    [Arguments("=AVERAGEIF(B1:B4,D1,C1:C4)", "1")]
    [Arguments("=MAXIFS(C1:C4,B1:B4,D1)", "1")]
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
    [Arguments("=XLOOKUP(D1,B1:B4,C1:C4)", "2")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,FALSE)", "2")]
    [Arguments("=VLOOKUP(D1,B1:C4,2,TRUE)", "2")]
    [Arguments("=COUNTIF(B1:B4,D1)", "2")]
    [Arguments("=SUMIF(B1:B4,D1,C1:C4)", "6")]
    [Arguments("=COUNTIFS(B1:B4,D1)", "2")]
    public async Task EmptyTextKey_RemainsDistinctFromAbsent(string formula, string expected) =>
        await Assert.That(Evaluate(formula, Key.FormulaEmpty, mixed: true)).IsEqualTo(expected);

    // Aspose 26.7.0 PLAIN/CSE: without a text candidate the old absent match 2 becomes #N/A (or the
    // supplied reverse-search fallback 0); with one at position 3, old forward results 2 become 3/4.
    [Test]
    [Arguments("=MATCH(D1,B1:B4,1)", "#N/A")]
    [Arguments("=MATCH(D1,B1:B4,-1)", "#N/A")]
    public async Task ApproximateMatch_EmptyTextWithoutMatch_IsNotAvailable(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, Key.ExplicitEmpty, mixed: false)).IsEqualTo(expected);

    [Test]
    [Arguments("=XMATCH(\"\",B1:B3,-1,1)")]
    [Arguments("=XMATCH(H10,B1:B3,-1,-1)")]
    [Arguments("=XMATCH(\"\",B1:B3,1,2)")]
    [Arguments("=XMATCH(H10,B1:B3,1,-2)")]
    [Arguments("=XLOOKUP(\"\",B1:B3,C1:C3,,-1,1)")]
    [Arguments("=XLOOKUP(H10,B1:B3,C1:C3,,-1,-1)")]
    [Arguments("=XLOOKUP(\"\",B1:B3,C1:C3,,1,2)")]
    [Arguments("=XLOOKUP(H10,B1:B3,C1:C3,,1,-2)")]
    public async Task ApproximateModernLookup_EmptyTextWithoutTextCandidate_IsNotAvailable(
        string formula
    ) =>
        // Aspose 26.7.0 PLAIN/CSE are #N/A; before the fix mode -1 returned position 3/result 4.
        await Assert
            .That(EvaluateExactEmptyText(formula, false, horizontal: false))
            .IsEqualTo("#N/A");

    // Aspose 26.7.0 PLAIN/CSE: LOOKUP's old zero/position-2 results become #N/A without exact text and
    // select the last exact-text candidate with it (empty text in two-argument form, result 4 in three).
    [Test]
    [Arguments("=COUNTIF(B1:B4,\"\")", "2")]
    [Arguments("=COUNTIF(B1:B4,\"=\")", "1")]
    [Arguments("=COUNTIF(B1:B4,\"<>\")", "3")]
    public async Task EmptyCriteria_DistinguishImplicitAndExplicitEquality(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, Key.Absent, mixed: true)).IsEqualTo(expected);

    // The values stay ROWS=2 and XLOOKUP=#N/A; only the volatile draw count changes from 2 to 1.
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

    // Aspose 26.7.0 PLAIN/CSE resolves each key to absent D1 before classifying it. The discriminating
    // C1:C4 = 1,2,4,8 fixture proves SUMIF selected B1: the four derived pins move 0 -> 1, not to B3's 4.
    [Test]
    [Arguments("=XMATCH(INDEX(D1:D1,1),B1:B4)", "3")]
    [Arguments("=XMATCH(OFFSET(D1,0,0),B1:B4)", "3")]
    [Arguments("=XMATCH(IF(TRUE,D1),B1:B4)", "3")]
    [Arguments("=XMATCH(EmptyCell,B1:B4)", "3")]
    [Arguments("=MATCH(INDEX(D1:D1,1),B1:B4,-1)", "1")]
    [Arguments("=MATCH(OFFSET(D1,0,0),B1:B4,-1)", "1")]
    [Arguments("=MATCH(IF(TRUE,D1),B1:B4,-1)", "1")]
    [Arguments("=MATCH(EmptyCell,B1:B4,-1)", "1")]
    [Arguments("=XLOOKUP(INDEX(D1:D1,1),B1:B4,C1:C4)", "4")]
    [Arguments("=XLOOKUP(OFFSET(D1,0,0),B1:B4,C1:C4)", "4")]
    [Arguments("=XLOOKUP(IF(TRUE,D1),B1:B4,C1:C4)", "4")]
    [Arguments("=XLOOKUP(EmptyCell,B1:B4,C1:C4)", "4")]
    [Arguments("=COUNTIF(B1:B4,INDEX(D1:D1,1))", "1")]
    [Arguments("=COUNTIF(B1:B4,OFFSET(D1,0,0))", "1")]
    [Arguments("=COUNTIF(B1:B4,IF(TRUE,D1))", "1")]
    [Arguments("=COUNTIF(B1:B4,EmptyCell)", "1")]
    [Arguments("=SUMIF(B1:B4,INDEX(D1:D1,1),C1:C4)", "1")]
    [Arguments("=SUMIF(B1:B4,OFFSET(D1,0,0),C1:C4)", "1")]
    [Arguments("=SUMIF(B1:B4,IF(TRUE,D1),C1:C4)", "1")]
    [Arguments("=SUMIF(B1:B4,EmptyCell,C1:C4)", "1")]
    public async Task DerivedAbsentKey_UsesTheConsumerFamilyRule(string formula, string expected) =>
        await Assert.That(Evaluate(formula, Key.Absent, mixed: true)).IsEqualTo(expected);

    [Test]
    [Arguments("LET(r,INDEX(D1:D1,1),r)")]
    [Arguments("LET(r,OFFSET(D1,0,0),r)")]
    [Arguments("LET(r,IF(TRUE,D1),r)")]
    [Arguments("LET(r,EmptyCell,r)")]
    [Arguments("LET(a,OFFSET(D1,0,0),LET(b,a,b))")]
    public async Task LetWrappedDerivedAbsentKey_UsesTheConsumerFamilyRule(string key)
    {
        await Assert.That(Evaluate($"=XMATCH({key},B1:B4)", Key.Absent, true)).IsEqualTo("3");
        await Assert.That(Evaluate($"=MATCH({key},B1:B4,-1)", Key.Absent, true)).IsEqualTo("1");
        await Assert
            .That(Evaluate($"=XLOOKUP({key},B1:B4,C1:C4)", Key.Absent, true))
            .IsEqualTo("4");
        await Assert.That(Evaluate($"=COUNTIF(B1:B4,{key})", Key.Absent, true)).IsEqualTo("1");
        await Assert.That(Evaluate($"=SUMIF(B1:B4,{key},C1:C4)", Key.Absent, true)).IsEqualTo("1");
    }

    [Test]
    [Arguments("=MATCH(H10,B1:B3,1)", MatchFixture.AbsentCandidate, "#N/A")]
    [Arguments("=MATCH(H10,B1:B3,-1)", MatchFixture.AbsentCandidate, "#N/A")]
    [Arguments("=MATCH(H10,B1:B3,1)", MatchFixture.OneText, "2")]
    [Arguments("=MATCH(H10,B1:B3,-1)", MatchFixture.OneText, "2")]
    [Arguments("=MATCH(H10,B1:B4,1)", MatchFixture.SeveralTexts, "3")]
    [Arguments("=MATCH(H10,B1:B4,-1)", MatchFixture.SeveralTexts, "2")]
    [Arguments("=MATCH(H10,B1:B2,1)", MatchFixture.AllAbsent, "#N/A")]
    [Arguments("=MATCH(H10,B1:B2,-1)", MatchFixture.AllAbsent, "#N/A")]
    [Arguments("=MATCH(H10,{0,\"\",5},1)", MatchFixture.AllAbsent, "2")]
    [Arguments("=MATCH(H10,{0,\"\",5},-1)", MatchFixture.AllAbsent, "2")]
    [Arguments("=MATCH(H10,{0,5},1)", MatchFixture.AllAbsent, "#N/A")]
    [Arguments("=MATCH(H10,{0,5},-1)", MatchFixture.AllAbsent, "#N/A")]
    public async Task ApproximateMatch_EmptyTextRequiresAnExactTextCandidate(
        string formula,
        MatchFixture fixture,
        string expected
    ) => await Assert.That(EvaluateApproximateMatch(formula, fixture)).IsEqualTo(expected);

    [Test]
    [Arguments("=MATCH(2,B1:B5,1)", false, "1")]
    [Arguments("=MATCH(4,B1:B5,1)", false, "3")]
    [Arguments("=MATCH(6,B1:B5,1)", false, "5")]
    [Arguments("=MATCH(4,B1:B5,-1)", true, "1")]
    [Arguments("=MATCH(2,B1:B5,-1)", true, "3")]
    [Arguments("=MATCH(0,B1:B5,-1)", true, "5")]
    public async Task ApproximateMatch_SkipsAbsentCells(
        string formula,
        bool descending,
        string expected
    ) =>
        await Assert
            .That(EvaluateApproximateMatchWithGaps(formula, descending))
            .IsEqualTo(expected);

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
    [Arguments("=XMATCH(LET(r,OFFSET(D1,TICK()*0,0),r),B1:B4)", "3")]
    [Arguments("=COUNTIF(B1:B4,LET(r,OFFSET(D1,TICK()*0,0),r))", "1")]
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

    [Test]
    [Arguments("=MATCH(\"\",B1:B3,0)", false, false, "#N/A")]
    [Arguments("=XMATCH(H10,B1:B3)", false, false, "#N/A")]
    [Arguments("=XLOOKUP(\"\",B1:B3,C1:C3)", false, false, "#N/A")]
    [Arguments("=VLOOKUP(H10,B1:C3,2,FALSE)", false, false, "#N/A")]
    [Arguments("=HLOOKUP(\"\",B1:D2,2,FALSE)", false, true, "#N/A")]
    [Arguments("=XLOOKUP(H10,B1:B3,C1:C3,0,0,-1)", false, false, "0")]
    [Arguments("=MATCH(H10,B1:B4,0)", true, false, "3")]
    [Arguments("=XMATCH(\"\",B1:B4)", true, false, "3")]
    [Arguments("=XLOOKUP(H10,B1:B4,C1:C4)", true, false, "4")]
    [Arguments("=VLOOKUP(\"\",B1:C4,2,FALSE)", true, false, "4")]
    [Arguments("=HLOOKUP(H10,B1:E2,2,FALSE)", true, true, "4")]
    [Arguments("=XLOOKUP(\"\",B1:B4,C1:C4,0,0,-1)", true, false, "4")]
    public async Task ExactEmptyText_MatchesOnlyTextCandidates(
        string formula,
        bool textCandidate,
        bool horizontal,
        string expected
    ) =>
        await Assert
            .That(EvaluateExactEmptyText(formula, textCandidate, horizontal))
            .IsEqualTo(expected);

    [Test]
    [Arguments("=LOOKUP(\"\",B1:B3)", false, "#N/A")]
    [Arguments("=LOOKUP(H10,B1:B3,C1:C3)", false, "#N/A")]
    [Arguments("=LOOKUP(\"\",B1:B4)", true, "\"\"")]
    [Arguments("=LOOKUP(H10,B1:B4,C1:C4)", true, "4")]
    [Arguments("=LOOKUP(\"\",{0,\"\",5})", true, "\"\"")]
    [Arguments("=LOOKUP(H10,{0,5})", true, "#N/A")]
    public async Task Lookup_EmptyText_UsesTheLastExactTextCandidate(
        string formula,
        bool textCandidate,
        string expected
    ) =>
        await Assert
            .That(EvaluateExactEmptyText(formula, textCandidate, horizontal: false))
            .IsEqualTo(expected);

    [Test]
    public async Task Lookup_EmptyText_SelectsTheLastOfSeveralTextCandidates()
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["B1"] = new NumberValue(0);
        main["B2"] = ExpressionParser.Parse("=\"\"", main);
        main["B3"] = ExpressionParser.Parse("=\"\"", main);
        main["B4"] = new NumberValue(5);
        main["C1"] = new NumberValue(1);
        main["C2"] = new NumberValue(2);
        main["C3"] = new NumberValue(4);
        main["C4"] = new NumberValue(8);
        main["AZ5000"] = ExpressionParser.Parse("=LOOKUP(\"\",B1:B4,C1:C4)", main);

        // Aspose 26.7.0 PLAIN/CSE returns 4; choosing the first exact text returns the old value 2.
        await Assert.That(Format(workbook.GetCellValue("Main", "AZ5000"))).IsEqualTo("4");
    }

    [Test]
    [Arguments("=LET(r,OFFSET(D1,TICK()*0,0,2,1),ROWS(r))", "2")]
    [Arguments("=XLOOKUP(OFFSET(D1,TICK()*0,0,2,1),B1:B4,C1:C4)", "#N/A")]
    public async Task MultiCellOffset_IsResolvedOnlyOnce(string formula, string expected)
    {
        var draws = 0;
        var workbook = CreateWorkbook(Key.Absent, mixed: true);
        workbook.RegisterFunction("TICK", (_, _) => ++draws);
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse(formula, main);

        await Assert.That(Format(workbook.GetCellValue("Main", "AZ5000"))).IsEqualTo(expected);
        await Assert.That(draws).IsEqualTo(1);
    }

    [Test]
    [Arguments("D1:D2", "#VALUE!")]
    [Arguments("D:D", "#VALUE!")]
    [Arguments("(D1,D2)", "#VALUE!")]
    [Arguments("OFFSET(D1,0,0,2,1)", "4")]
    [Arguments("INDEX(D1:E2,0,1)", "4")]
    [Arguments("LET(r,D1:D2,r)", "4")]
    [Arguments("KeyRange", "4")]
    public async Task MultiCellLookupKey_UsesOneRuleAcrossReferenceKinds(
        string key,
        string expected
    )
    {
        var workbook = CreateWorkbook(Key.FormulaEmpty, mixed: true);
        var main = workbook.Sheets["Main"];
        main["D1"] = new NumberValue(5);
        main["D2"] = new NumberValue(0);
        workbook.DefineName("KeyRange", "Main!D1:D2");
        main["AZ5000"] = ExpressionParser.Parse($"=MATCH({key},B1:B4,1)", main);

        // Item 75 lifting remains deferred. Aspose 26.7.0 CSE is 4 for every row; the direct-reference
        // rows intentionally retain their pre-Phase-5b #VALUE! boundary while derived references retain 4.
        await Assert.That(Format(workbook.GetCellValue("Main", "AZ5000"))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=MATCH(\"\",{0,\"\",5},0)", "2")]
    [Arguments("=MATCH(\"\",{0,5},0)", "#N/A")]
    [Arguments("=XMATCH(\"\",{0,\"\",5})", "2")]
    [Arguments("=XMATCH(\"\",{0,5})", "#N/A")]
    [Arguments("=XLOOKUP(\"\",{0,\"\",5},{1,4,8})", "4")]
    [Arguments("=XLOOKUP(\"\",{0,5},{1,8})", "#N/A")]
    [Arguments("=XMATCH(\"\",B1:B3,2)", "2")]
    [Arguments("=MATCH(2,{1,2,3},0)", "2")]
    [Arguments("=MATCH(\"A\",{\"a\",\"b\"},0)", "1")]
    public async Task ExactEmptyText_GuardsStayUnchanged(string formula, string expected)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["B1"] = new NumberValue(0);
        main["B2"] = ExpressionParser.Parse("=\"\"", main);
        main["B3"] = new NumberValue(5);
        main["AZ5000"] = ExpressionParser.Parse(formula, main);

        await Assert.That(Format(workbook.GetCellValue("Main", "AZ5000"))).IsEqualTo(expected);
    }

    // Aspose.Cells 26.7.0 PLAIN/CSE agree on every row. Before the fix, wildcard mode treated the
    // absent key as numeric zero or empty text; these 22 rows instead require an actually absent cell.
    [Test]
    [Arguments(MatchFixture.ZeroAbsentText, "=XMATCH(D1,A1:A3,2)", "2")]
    [Arguments(MatchFixture.ZeroAbsentText, "=XMATCH(D1,A1:A3,2,-1)", "2")]
    [Arguments(MatchFixture.ZeroAbsentText, "=XLOOKUP(D1,A1:A3,{1;2;4},,2)", "2")]
    [Arguments(MatchFixture.ZeroAbsentText, "=XLOOKUP(D1,A1:A3,{1;2;4},,2,-1)", "2")]
    [Arguments(MatchFixture.ZeroTextAbsent, "=XMATCH(D1,A1:A3,2)", "3")]
    [Arguments(MatchFixture.ZeroTextAbsent, "=XMATCH(D1,A1:A3,2,-1)", "3")]
    [Arguments(MatchFixture.ZeroTextAbsent, "=XLOOKUP(D1,A1:A3,{1;2;4},,2)", "4")]
    [Arguments(MatchFixture.ZeroTextAbsent, "=XLOOKUP(D1,A1:A3,{1;2;4},,2,-1)", "4")]
    [Arguments(MatchFixture.AllAbsent, "=XMATCH(D1,A1:A2,2,-1)", "1")]
    [Arguments(MatchFixture.AllAbsent, "=XLOOKUP(D1,A1:A2,{1;2},,2,-1)", "1")]
    [Arguments(MatchFixture.ZeroFive, "=XMATCH(D1,A1:A2,2)", "#N/A")]
    [Arguments(MatchFixture.ZeroFive, "=XMATCH(D1,A1:A2,2,-1)", "#N/A")]
    [Arguments(MatchFixture.ZeroFive, "=XLOOKUP(D1,A1:A2,{1;2},,2)", "#N/A")]
    [Arguments(MatchFixture.ZeroFive, "=XLOOKUP(D1,A1:A2,{1;2},,2,-1)", "#N/A")]
    [Arguments(MatchFixture.TextAbsentZero, "=XMATCH(D1,A1:A3,2)", "2")]
    [Arguments(MatchFixture.TextAbsentZero, "=XMATCH(D1,A1:A3,2,-1)", "2")]
    [Arguments(MatchFixture.TextAbsentZero, "=XLOOKUP(D1,A1:A3,{1;2;4},,2)", "2")]
    [Arguments(MatchFixture.TextAbsentZero, "=XLOOKUP(D1,A1:A3,{1;2;4},,2,-1)", "2")]
    [Arguments(MatchFixture.Array, "=XMATCH(D1,{0,\"\",5},2)", "#N/A")]
    [Arguments(MatchFixture.Array, "=XMATCH(D1,{0,\"\",5},2,-1)", "#N/A")]
    [Arguments(MatchFixture.Array, "=XLOOKUP(D1,{0,\"\",5},{1,2,4},,2)", "#N/A")]
    [Arguments(MatchFixture.Array, "=XLOOKUP(D1,{0,\"\",5},{1,2,4},,2,-1)", "#N/A")]
    public async Task WildcardMode_AbsentKey_MatchesOnlyAbsentCells(
        MatchFixture fixture,
        string formula,
        string expected
    ) => await Assert.That(EvaluateWildcardAbsent(formula, fixture)).IsEqualTo(expected);

    // Aspose.Cells 26.7.0 PLAIN/CSE agree: wildcard-mode numeric zero matches numeric zero only.
    // Before the fix these reverse or blank-only rows selected an absent/formula-empty candidate.
    [Test]
    [Arguments(MatchFixture.ZeroAbsentText, "=XMATCH(0,A1:A3,2,-1)", "1")]
    [Arguments(MatchFixture.ZeroAbsentText, "=XLOOKUP(0,A1:A3,{1;2;4},,2,-1)", "1")]
    [Arguments(MatchFixture.ZeroTextAbsent, "=XMATCH(0,A1:A3,2,-1)", "1")]
    [Arguments(MatchFixture.ZeroTextAbsent, "=XLOOKUP(0,A1:A3,{1;2;4},,2,-1)", "1")]
    [Arguments(MatchFixture.AllAbsent, "=XMATCH(0,A1:A2,2)", "#N/A")]
    [Arguments(MatchFixture.AllAbsent, "=XMATCH(0,A1:A2,2,-1)", "#N/A")]
    [Arguments(MatchFixture.AllAbsent, "=XLOOKUP(0,A1:A2,{1;2},,2)", "#N/A")]
    [Arguments(MatchFixture.AllAbsent, "=XLOOKUP(0,A1:A2,{1;2},,2,-1)", "#N/A")]
    [Arguments(MatchFixture.TextAbsentZero, "=XMATCH(0,A1:A3,2)", "3")]
    [Arguments(MatchFixture.TextAbsentZero, "=XLOOKUP(0,A1:A3,{1;2;4},,2)", "4")]
    public async Task WildcardMode_ZeroKey_MatchesOnlyNumericZero(
        MatchFixture fixture,
        string formula,
        string expected
    ) => await Assert.That(EvaluateWildcardAbsent(formula, fixture)).IsEqualTo(expected);

    [Test]
    [Arguments("=XMATCH(\"\",A1:A2,2)")]
    [Arguments("=XMATCH(H10,A1:A2,2,-1)")]
    [Arguments("=XLOOKUP(IF(TRUE,\"\"),A1:A2,{1;2},,2)")]
    public async Task WildcardMode_EmptyText_DoesNotMatchAbsentCells(string formula) =>
        await Assert
            .That(EvaluateWildcardAbsent(formula, MatchFixture.AllAbsent))
            .IsEqualTo("#N/A");

    private static string Evaluate(string formula, Key key, bool mixed)
    {
        var workbook = CreateWorkbook(key, mixed);
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return Format(workbook.GetCellValue("Main", "AZ5000"));
    }

    private static string EvaluateExactEmptyText(
        string formula,
        bool textCandidate,
        bool horizontal
    )
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        if (horizontal)
        {
            main["B1"] = new NumberValue(0);
            if (textCandidate)
            {
                main["D1"] = ExpressionParser.Parse("=\"\"", main);
            }
            main[textCandidate ? "E1" : "D1"] = new NumberValue(5);
            for (var column = 0; column < 4; column++)
            {
                main[$"{(char)('B' + column)}2"] = new NumberValue(1 << column);
            }
        }
        else
        {
            main["B1"] = new NumberValue(0);
            if (textCandidate)
            {
                main["B3"] = ExpressionParser.Parse("=\"\"", main);
            }
            main[textCandidate ? "B4" : "B3"] = new NumberValue(5);
            for (var row = 1; row <= 4; row++)
            {
                main[$"C{row}"] = new NumberValue(1 << (row - 1));
            }
        }
        main["H10"] = new Danfma.MySheet.Expressions.StringValue(string.Empty);
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
            main[$"C{row}"] = new NumberValue(1 << (row - 1));
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

    private static string EvaluateApproximateMatch(string formula, MatchFixture fixture)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["H10"] = new Danfma.MySheet.Expressions.StringValue(string.Empty);
        if (fixture is not MatchFixture.AllAbsent)
        {
            main["B1"] = new NumberValue(0);
            main[fixture is MatchFixture.SeveralTexts ? "B4" : "B3"] = new NumberValue(5);
        }
        if (fixture is MatchFixture.OneText or MatchFixture.SeveralTexts)
        {
            main["B2"] = ExpressionParser.Parse("=\"\"", main);
        }
        if (fixture is MatchFixture.SeveralTexts)
        {
            main["B3"] = ExpressionParser.Parse("=\"\"", main);
        }
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return Format(workbook.GetCellValue("Main", "AZ5000"));
    }

    private static string EvaluateWildcardAbsent(string formula, MatchFixture fixture)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["H10"] = new Danfma.MySheet.Expressions.StringValue(string.Empty);
        if (fixture == MatchFixture.ZeroAbsentText)
        {
            main["A1"] = new NumberValue(0);
            main["A3"] = ExpressionParser.Parse("=\"\"", main);
        }
        else if (fixture == MatchFixture.ZeroTextAbsent)
        {
            main["A1"] = new NumberValue(0);
            main["A2"] = ExpressionParser.Parse("=\"\"", main);
        }
        else if (fixture == MatchFixture.ZeroFive)
        {
            main["A1"] = new NumberValue(0);
            main["A2"] = new NumberValue(5);
        }
        else if (fixture == MatchFixture.TextAbsentZero)
        {
            main["A1"] = ExpressionParser.Parse("=\"\"", main);
            main["A3"] = new NumberValue(0);
        }
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return Format(workbook.GetCellValue("Main", "AZ5000"));
    }

    private static string EvaluateApproximateMatchWithGaps(string formula, bool descending)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["B1"] = new NumberValue(descending ? 5 : 1);
        main["B3"] = new NumberValue(3);
        main["B5"] = new NumberValue(descending ? 1 : 5);
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

    public enum MatchFixture
    {
        AbsentCandidate,
        OneText,
        SeveralTexts,
        AllAbsent,
        ZeroAbsentText,
        ZeroTextAbsent,
        ZeroFive,
        TextAbsentZero,
        Array,
    }
}
