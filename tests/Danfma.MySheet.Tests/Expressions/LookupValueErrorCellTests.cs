using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The lookup VALUE slot over a single CELL holding an error. Fixture = the controller's divergence probe:
/// Main!A1 = <c>=1/0</c>, B1:B3 = 1, 2, 3, C1:C3 = 10, 20, 30, D1 = 5, A5:A7 = 5, 0, 9, and the defined name
/// <c>ErrCell</c> = <c>Main!$A$1</c>; the formula sits at Main!AZ5000 and is read back through the cell.
/// <para>
/// Measured on Aspose.Cells 26.6.0 AND 26.7.0 (the oracle migration, 2026-09-14), one formula per workbook,
/// PLAIN and array-entered: the two versions and the two entry modes agree on every row. The lookup value's
/// error leads every lookup — <c>MATCH(A1,A1)</c> and every match type, <c>XMATCH</c>, <c>XLOOKUP</c> even with
/// an <c>if_not_found</c>, <c>VLOOKUP</c>, <c>HLOOKUP</c>, <c>LOOKUP</c> — whether the cell is written directly or
/// reached through a name.
/// </para>
/// </summary>
public class LookupValueErrorCellTests
{
    private static string InCell(string formula)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = ExpressionParser.Parse("=1/0", main);
        main["B1"] = new NumberValue(1);
        main["B2"] = new NumberValue(2);
        main["B3"] = new NumberValue(3);
        main["C1"] = new NumberValue(10);
        main["C2"] = new NumberValue(20);
        main["C3"] = new NumberValue(30);
        main["D1"] = new NumberValue(5);
        main["A5"] = new NumberValue(5);
        main["A6"] = new NumberValue(0);
        main["A7"] = new NumberValue(9);
        workbook.DefineName("ErrCell", "Main!$A$1");
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

    // The regression and its neighbours: MySheet answered #N/A (or a position) where both oracle versions
    // answer the cell's #DIV/0!. Before this fix (branch 4b82423) → oracle:
    //   MATCH(A1,A1)            #N/A → #DIV/0!   (main b778ab7 answered #DIV/0!: T1's value-slot guard lost it)
    //   MATCH(A1,A1,1)          #N/A → #DIV/0!
    //   MATCH(A1,A1,-1)         #N/A → #DIV/0!
    //   MATCH(A1,A1,0)          #N/A → #DIV/0!   (main #N/A too — "Bug 8")
    //   ISNA(MATCH(A1,A1,0))    TRUE → FALSE
    //   MATCH(A1,B1:B3,0)       #N/A → #DIV/0!;  MATCH(A1,B1:B3) 3 → #DIV/0!;  MATCH(A1,B1:B3,-1) #N/A → #DIV/0!
    //   MATCH(ErrCell,B1:B3,0)  #N/A → #DIV/0!;  MATCH(ErrCell,B1:B3) 3 → #DIV/0!
    //   XMATCH(A1,B1:B3), XMATCH(A1,A1), XMATCH(ErrCell,B1:B3)                     #N/A → #DIV/0!
    //   XLOOKUP(A1,B1:B3,C1:C3) #N/A, with "nf" "nf", XLOOKUP(ErrCell,B1:B3,C1:C3) #N/A → #DIV/0!
    [Test]
    [Arguments("=MATCH(A1,A1)", "#DIV/0!")]
    [Arguments("=MATCH(A1,A1,1)", "#DIV/0!")]
    [Arguments("=MATCH(A1,A1,-1)", "#DIV/0!")]
    [Arguments("=MATCH(A1,A1,0)", "#DIV/0!")]
    [Arguments("=ISNA(MATCH(A1,A1,0))", "FALSE")]
    [Arguments("=MATCH(A1,B1:B3,0)", "#DIV/0!")]
    [Arguments("=MATCH(A1,B1:B3)", "#DIV/0!")]
    [Arguments("=MATCH(A1,B1:B3,-1)", "#DIV/0!")]
    [Arguments("=MATCH(ErrCell,B1:B3,0)", "#DIV/0!")]
    [Arguments("=MATCH(ErrCell,B1:B3)", "#DIV/0!")]
    [Arguments("=XMATCH(A1,B1:B3)", "#DIV/0!")]
    [Arguments("=XMATCH(A1,A1)", "#DIV/0!")]
    [Arguments("=XMATCH(ErrCell,B1:B3)", "#DIV/0!")]
    [Arguments("=XLOOKUP(A1,B1:B3,C1:C3)", "#DIV/0!")]
    [Arguments("=XLOOKUP(A1,B1:B3,C1:C3,\"nf\")", "#DIV/0!")]
    [Arguments("=XLOOKUP(ErrCell,B1:B3,C1:C3)", "#DIV/0!")]
    public async Task ALookupValueCellHoldingAnError_LeadsTheLookup(string formula, string expected)
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // The rows that already agreed before the fix and must keep agreeing: a computed value slot (1/0,
    // INDIRECT and INDEX evaluate to the error themselves) and the three lookups that check the value's
    // error unconditionally. Oracle 26.6.0 = 26.7.0, both modes: #DIV/0! for all of them.
    [Test]
    [Arguments("=MATCH(1/0,1/0,0)", "#DIV/0!")]
    [Arguments("=MATCH(INDIRECT(\"A1\"),B1:B3,0)", "#DIV/0!")]
    [Arguments("=MATCH(INDEX(A1:B1,1,1),B1:B3,0)", "#DIV/0!")]
    [Arguments("=VLOOKUP(A1,B1:C3,2,FALSE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(A1,B1:C3,2)", "#DIV/0!")]
    [Arguments("=HLOOKUP(A1,B1:D1,1,FALSE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(A1,B1:D1,1)", "#DIV/0!")]
    [Arguments("=LOOKUP(A1,B1:B3)", "#DIV/0!")]
    [Arguments("=LOOKUP(A1,B1:B3,C1:C3)", "#DIV/0!")]
    public async Task TheLookupsThatAlreadyPropagated_KeepPropagating(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // The guard the value-slot rule must NOT cross: a RANGE in the value slot collapses to #VALUE! (a range has
    // no scalar value), which is an artifact, not the lookup value's own error — the reason the shared guard
    // exists. MySheet does not lift a Consumes function over its scalar slot (ElementwiseLiftingTests' known
    // divergence), so these stay where they were. Oracle 26.6.0 = 26.7.0, PLAIN / CSE:
    //   SUM(MATCH(A5:A7,A5:A7,0)) #VALUE! / 6, MATCH(A5:A7,A5:A7,0) #VALUE! / 1, SUM(MATCH(B1:B3,B1:B3)) #VALUE! / 6.
    [Test]
    [Arguments("=SUM(MATCH(A5:A7,A5:A7,0))", "#N/A")]
    [Arguments("=MATCH(A5:A7,A5:A7,0)", "#N/A")]
    [Arguments("=SUM(MATCH(B1:B3,B1:B3))", "3")]
    public async Task ARangeInTheValueSlot_KeepsItsCollapseArtifact(string formula, string expected)
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }
}
