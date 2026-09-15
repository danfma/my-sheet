using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Sweep item 34(b) / the MATCH fix, final-review finding I2 (Important). RULING (the fix wave's binding
/// brief): pin the move, register a sweep item for the controller, NO CODE CHANGE.
/// <para>
/// <see cref="Danfma.MySheet.Expressions.Lookup.ReferencePosition.TryUnresolvedError"/> has no
/// "collapse-artifact" guard the way <c>PositionalRange.IsOwnSlotError</c> does for the criteria family: a
/// COMPUTED lookup ARRAY — <c>A1:A3*1</c>, <c>Rng*1</c>, <c>CHOOSE(1,A1:A3)*1</c>, anything that is not a
/// bare reference NODE, so <c>NamedReferences.TryResolveReference</c> cannot see it — falls through to
/// evaluating the node directly, and a lifted range's scalar <c>Evaluate</c> is <c>#VALUE!</c> (the same
/// collapse artifact sweep item 32 removed from the C1 branch shape, <c>RangeReference.Evaluate</c>'s
/// unconditional rule). <c>TryUnresolvedError</c> then reads that <c>#VALUE!</c> as the argument's "own
/// error" and propagates it — silently turning main's <c>#N/A</c>/<c>#REF!</c> into <c>#VALUE!</c> for
/// resolving consumers routed through that shared site (MATCH, XMATCH, LOOKUP and OFFSET). VLOOKUP and
/// HLOOKUP now choose an array-valued table through <c>ArrayEvaluation.TryStream</c> before resolution.
/// </para>
/// <para>
/// This is NOT a regression against the oracle — <c>#VALUE!</c> equals Aspose.Cells' PLAIN-entry answer on
/// every row below — but it MOVED SILENTLY (no pin, no commit-body line, across T1's <c>a542f56</c> and the
/// MATCH fix's <c>5975c63</c>) and the branch does not reach the oracle's CSE (array-entered) answer
/// either. The controller registers "a lookup array slot over a lifted computation (oracle CSE lifts)" as a
/// NEW sweep item. This class retains those pins and records the table consumers that have since moved to
/// the CSE route.
/// </para>
/// <para>
/// Fixture: Main!A1:A3 = 5, 0, 9; B1:B3 = 1, 2, 3; the defined name <c>Rng</c> = <c>Main!$A$1:$A$3</c>;
/// <c>Data!Tabela1</c> = A1:C4 (header + 3 rows), the Valor column (B) = 10, 20, 30 — the same table shape
/// <c>MiniCseConsumerTests.TableGrid</c> uses. Measured on Aspose.Cells 26.7.0 (= 26.6.0), one formula per
/// workbook, PLAIN and array-entered (CSE): main = <c>a02ed5d</c>'s library source (pre-sweep, extracted to
/// scratch); branch = this worktree.
/// </para>
/// </summary>
public class LiftedLookupArraySlotTests
{
    private static string On(string formula)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(5);
        main["A2"] = new NumberValue(0);
        main["A3"] = new NumberValue(9);
        main["B1"] = new NumberValue(1);
        main["B2"] = new NumberValue(2);
        main["B3"] = new NumberValue(3);
        workbook.DefineName("Rng", "Main!$A$1:$A$3");

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

        if (value.TryGetBoolean(out var boolean))
        {
            return boolean ? "TRUE" : "FALSE";
        }

        return value.TryGetText(out var text) ? $"\"{text}\"" : value.Kind.ToString();
    }

    // MATCH's VALUE slot over a lifted array stays on its scalar path. XMATCH now applies its general computed-
    // array shape/scan rule, changing #VALUE! -> 1 to match CSE (PLAIN remains #VALUE!). The MATCH(9,...+0)
    // CSE result is 3; that separate consumer keeps the previously measured scalar behavior.
    [Test]
    [Arguments("=MATCH(5,A1:A3*1,0)", "#VALUE!")]
    [Arguments("=MATCH(9,A1:A3+0,0)", "#VALUE!")]
    [Arguments("=MATCH(5,Rng*1,0)", "#VALUE!")]
    [Arguments("=MATCH(5,CHOOSE(1,A1:A3)*1,0)", "#VALUE!")]
    [Arguments("=XMATCH(5,A1:A3*1)", "1")]
    public async Task ALiftedComputedArray_InTheMatchValueSlot_PropagatesTheCollapseArtifact(
        string formula,
        string expected
    )
    {
        await Assert.That(On(formula)).IsEqualTo(expected);
    }

    // The structured-reference twin, over a 3-row Tabela1[Valor] = 10, 20, 30: branch #VALUE!, main #N/A,
    // oracle PLAIN #VALUE! / CSE 3 (30 is the third row).
    [Test]
    public async Task ALiftedStructuredReference_InTheMatchValueSlot_PropagatesTheCollapseArtifact()
    {
        await Assert.That(On("=MATCH(30,Tabela1[Valor]*1,0)")).IsEqualTo("#VALUE!");
    }

    // LOOKUP's vector-form array slot now materializes the full computed vector. The old #VALUE! collapse
    // becomes 5, matching the vector's last value <= 5; Aspose PLAIN/CSE also consume this route as a vector.
    [Test]
    public async Task ALiftedComputedArray_InTheLookupVectorSlot_IsMaterialized()
    {
        await Assert.That(On("=LOOKUP(5,A1:A3*1)")).IsEqualTo("5");
    }

    [Test]
    [Arguments("=LOOKUP(2,{1,2,3}*TICK()^0)", "2")]
    [Arguments("=LOOKUP(2,{1,2,3}*TICK()^0,{10,20,30})", "20")]
    [Arguments("=LOOKUP(2,{1,2,3},{10,20,30}*TICK()^0)", "20")]
    [Arguments("=LOOKUP(2,B1:B4,{10,20,30,40}*TICK()^0)", "20")]
    [Arguments("=LOOKUP(2,B1:B4*TICK()^0,C1:C4)", "20")]
    [Arguments("=LOOKUP(2,B1:B4,C1:C4*TICK()^0)", "20")]
    [Arguments("=LOOKUP(TICK()*0+2,B1:B4,C1:C4)", "20")]
    [Arguments("=LOOKUP(2,IF(TICK()>0,NoSuch,B1:B4))", "#NAME?")]
    public async Task Lookup_MaterializesEachVolatileVectorOnce(string formula, string expected)
    {
        var draws = 0;
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        for (var row = 1; row <= 4; row++)
        {
            main[$"B{row}"] = new NumberValue(row);
            main[$"C{row}"] = new NumberValue(row * 10);
        }

        workbook.RegisterFunction("TICK", (_, _) => ++draws);
        main["AZ5000"] = ExpressionParser.Parse(formula, main);

        await Assert.That(OnCell(workbook)).IsEqualTo(expected);
        await Assert.That(draws).IsEqualTo(1);
    }

    [Test]
    [Arguments("=LOOKUP(1,{1,2}/0)", "#N/A")]
    [Arguments("=LOOKUP(2,1/(B1:B4=2),C1:C4)", "20")]
    [Arguments("=LOOKUP(2,{1,2}/{1,0},{10,20})", "10")]
    [Arguments("=LOOKUP(3,{1,#N/A,3},{10,20,30})", "30")]
    [Arguments("=LOOKUP(1,{1,2}/0,{10,20})", "#N/A")]
    [Arguments("=LOOKUP(2,B1:B4/(B1:B4<>3),C1:C4)", "20")]
    public async Task Lookup_SkipsErrorElementsInAComputedLookupVector(
        string formula,
        string expected
    )
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        for (var row = 1; row <= 4; row++)
        {
            main[$"B{row}"] = new NumberValue(row);
            main[$"C{row}"] = new NumberValue(row * 10);
        }

        main["AZ5000"] = ExpressionParser.Parse(formula, main);

        // Aspose.Cells 26.7.0 PLAIN/CSE agree on every expected value above. Before computed-vector
        // materialization, rows 1/2/4/5/6 were #DIV/0!, #VALUE!, 10, #DIV/0!, #VALUE!, respectively.
        await Assert.That(OnCell(workbook)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=LOOKUP(1,1/0)", "#DIV/0!")]
    [Arguments("=LOOKUP(1,{1}/0)", "#N/A")]
    [Arguments("=LOOKUP(1,A1/0)", "#N/A")]
    [Arguments("=LOOKUP(1,A1:A1/0)", "#N/A")]
    [Arguments("=LOOKUP(2,IF(FALSE,NoSuch,B1:B4))", "2")]
    [Arguments("=LOOKUP(2,CHOOSE(1,NoSuch,B1:B4))", "#NAME?")]
    [Arguments("=LOOKUP(1,NoSheet!A1*1)", "#REF!")]
    [Arguments("=LOOKUP(1,(NoSheet!A1:A3)*{1;1;1})", "#N/A")]
    [Arguments("=LOOKUP(1,IF(TRUE,NoSheet!A1:A3,A1:A3))", "#REF!")]
    [Arguments("=LOOKUP(2,IF(TRUE,NoSuch,B1:B4))", "#NAME?")]
    public async Task Lookup_DistinguishesScalarVectorErrorsFromArrayElementErrors(
        string formula,
        string expected
    ) =>
        // Aspose 26.7.0 PLAIN/CSE agree except for literal 1/0 (#DIV/0! / #N/A); that pre-existing
        // PLAIN-compatible row remains #DIV/0!, while the red rows move to the agreed values above.
        await Assert.That(OnLookupFixture(formula)).IsEqualTo(expected);

    [Test]
    [Arguments("=LOOKUP(2,B1:B4,IF(TRUE,NoSuch,C1:C4))", "#N/A")]
    [Arguments("=LOOKUP(2,B1:B4,NoSheet!C1:C4*1)", "#REF!")]
    public async Task Lookup_ResultVectorPreservesItsMeasuredErrorSemantics(
        string formula,
        string expected
    ) => await Assert.That(OnLookupFixture(formula)).IsEqualTo(expected);

    // VLOOKUP/HLOOKUP now consume the lifted table through the same TryStream route as every array-valued
    // table, following the oracle CSE values. Before that shared route they returned #VALUE!; OFFSET keeps
    // its existing scalar-collapse behavior and is not part of the lookup-table change.
    [Test]
    [Arguments("=VLOOKUP(5,A1:B3*1,2,FALSE)", "1")]
    [Arguments("=HLOOKUP(5,A1:B3*1,2,FALSE)", "0")]
    [Arguments("=OFFSET(A1:A3*1,0,0)", "#VALUE!")]
    public async Task ALiftedComputedArray_UsesTheConsumerSpecificRoute(
        string formula,
        string expected
    )
    {
        await Assert.That(On(formula)).IsEqualTo(expected);
    }

    // The recorded inconsistency: the SAME lifted computation, captured through a LET binding first
    // (ArrayBindings.Capture resolves a range-eligible expression to its cells, not a scalar collapse), does
    // NOT hit TryUnresolvedError's blind spot — MATCH finds 5 at position 1. Sibling forms of the identical
    // shape answering differently (#VALUE! bare, 1 through LET) is the inconsistency I2 records beside the
    // rows above; not fixed here either.
    [Test]
    public async Task TheSameLiftedComputation_ThroughALetBinding_FindsThePosition_TheRecordedInconsistency()
    {
        await Assert.That(On("=MATCH(5,LET(x,A1:A3*1,x),0)")).IsEqualTo("1");
    }

    private static string OnCell(Workbook workbook)
    {
        var value = workbook.GetCellValue("Main", "AZ5000");
        if (value.TryGetError(out var error))
        {
            return error.ToString();
        }

        return value.TryGetNumber(out var number)
            ? number.ToString(CultureInfo.InvariantCulture)
            : value.AsString() ?? value.Kind.ToString();
    }

    private static string OnLookupFixture(string formula)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        for (var row = 1; row <= 4; row++)
        {
            main[$"A{row}"] = new NumberValue(row);
            main[$"B{row}"] = new NumberValue(row);
            main[$"C{row}"] = new NumberValue(row * 10);
        }

        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        return OnCell(workbook);
    }
}
