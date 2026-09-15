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

    // MATCH/XMATCH's VALUE slot over a lifted array. Branch #VALUE! (this class's own measurement); main
    // #N/A (a02ed5d, measured); oracle PLAIN #VALUE! / CSE 1, except MATCH(9,A1:A3+0,0) whose CSE is 3 (the
    // "+0" row lifts to a different position than the "*1" rows on the oracle — recorded exactly as Fable's
    // table has it, not reconciled here).
    [Test]
    [Arguments("=MATCH(5,A1:A3*1,0)", "#VALUE!")]
    [Arguments("=MATCH(9,A1:A3+0,0)", "#VALUE!")]
    [Arguments("=MATCH(5,Rng*1,0)", "#VALUE!")]
    [Arguments("=MATCH(5,CHOOSE(1,A1:A3)*1,0)", "#VALUE!")]
    [Arguments("=XMATCH(5,A1:A3*1)", "#VALUE!")]
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

    // LOOKUP's vector-form array slot: branch #VALUE!, main #N/A, oracle 0 / 0 — BOTH oracle columns agree
    // (unlike the MATCH rows above), so this row is the sharpest illustration that #VALUE! is neither
    // oracle answer, only PLAIN's for the others.
    [Test]
    public async Task ALiftedComputedArray_InTheLookupVectorSlot_PropagatesTheCollapseArtifact()
    {
        await Assert.That(On("=LOOKUP(5,A1:A3*1)")).IsEqualTo("#VALUE!");
    }

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
}
