using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// The lookup VALUE slot over a single CELL holding an error. Fixture = the controller's divergence probe:
/// Main!A1 = <c>=1/0</c>, B1:B3 = 1, 2, 3, C1:C3 = 10, 20, 30, D1 = 5, A5:A7 = 5, 0, 9, and the defined names
/// <c>ErrCell</c> = <c>Main!$A$1</c> and <c>ErrRange</c> = <c>Main!$A$1:$A$1</c> (finding I4's 1x1-range
/// shape); the formula sits at Main!AZ5000 and is read back through the cell.
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
        main["Z1"] = ExpressionParser.Parse("=1/0", main);
        workbook.DefineName("ErrCell", "Main!$A$1");
        workbook.DefineName("ErrRange", "Main!$A$1:$A$1");
        workbook.DefineName("ErrLookup", "Main!$Z$1");
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
    // exists. Final-review fix wave, finding I1: the comment that used to sit here ("these stay where they
    // were") was FALSE — MEASURED false against main (`#VALUE!`), and the branch (`5975c63`) had silently
    // moved `SUM(MATCH(B1:B3,B1:B3))` from `#VALUE!` to `3` (a POSITION) without a pin or a commit-body line,
    // the exact "never silently change an expected value" rule the sweep is built on. The EXACT path
    // (matchType 0) still "stays where it were": IsLookupValueError's narrower guard is the ONLY one on that
    // path, so `MATCH(A5:A7,A5:A7,0)`'s range-collapse `#VALUE!` is not "the lookup value's own error" and
    // the scan proceeds to its ordinary not-found `#N/A`. The APPROXIMATE path (the third row, matchType
    // omitted -> 1) now restores main's unconditional rule instead (Match.cs, right after the exact-path
    // block): the lookup value's error leads UNCONDITIONALLY there, so `MATCH(B1:B3,B1:B3)` is `#VALUE!`
    // directly and `SUM` of it propagates the same `#VALUE!` — branch 3 -> #VALUE!, both main's and the
    // oracle PLAIN's answer. Oracle 26.7.0 = 26.6.0, PLAIN / CSE:
    //   SUM(MATCH(A5:A7,A5:A7,0)) #VALUE! / 6, MATCH(A5:A7,A5:A7,0) #VALUE! / 1, SUM(MATCH(B1:B3,B1:B3)) #VALUE! / 6.
    [Test]
    [Arguments("=SUM(MATCH(A5:A7,A5:A7,0))", "#N/A")]
    [Arguments("=MATCH(A5:A7,A5:A7,0)", "#N/A")]
    [Arguments("=SUM(MATCH(B1:B3,B1:B3))", "#VALUE!")]
    public async Task ARangeInTheValueSlot_KeepsItsCollapseArtifact(string formula, string expected)
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // I1's other approximate-path rows, measured on Aspose.Cells 26.7.0 (= 26.6.0), PLAIN / CSE: the lookup
    // value's error leads UNCONDITIONALLY on this path, so a range-collapse #VALUE! propagates exactly like a
    // single cell's own error does (ALookupValueCellHoldingAnError_LeadsTheLookup above) instead of scanning
    // for a literal match. `MATCH(A5:A7,A5:A7)` (matchType omitted, approximate) moves the same way as the
    // B1:B3 row; both are new pins (T1/the MATCH fix never measured the approximate-path range case).
    [Test]
    [Arguments("=MATCH(B1:B3,B1:B3)", "#VALUE!")]
    [Arguments("=MATCH(A5:A7,A5:A7)", "#VALUE!")]
    public async Task ARangeInTheValueSlot_OnTheApproximatePath_PropagatesUnconditionally(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Final-review fix wave, finding I4 (Important, pre-existing on main). ReferencePosition.cs used to test
    // `reference is CellReference` only; A1:A1 and a name bound to $A$1:$A$1 (ErrRange below) both resolve to
    // a RangeReference instead — a 1x1 one, which the value-slot rule's own doc ("the argument denotes a
    // single CELL") already covers but the code did not. The widened IsLookupValueError reads the ONE cell
    // directly (RangeReference.CellComputedValueAt) instead of scanning for the RangeReference.Evaluate
    // collapse artifact (#VALUE!) that is never in the array — before this fix, #N/A on the branch AND on
    // main (both engines shared the same narrow test). Oracle 26.7.0 = 26.6.0, both modes: #DIV/0! for all
    // four; MATCH(A1:A1,B1:B3,0) is on the EXACT path (matchType 0), the other three default/approximate.
    [Test]
    [Arguments("=MATCH(A1:A1,B1:B3,0)", "#DIV/0!")]
    [Arguments("=MATCH(A1:A1,B1:B3)", "#DIV/0!")]
    [Arguments("=XMATCH(A1:A1,B1:B3)", "#DIV/0!")]
    [Arguments("=XLOOKUP(A1:A1,B1:B3,C1:C3)", "#DIV/0!")]
    [Arguments("=MATCH(ErrRange,B1:B3,0)", "#DIV/0!")]
    public async Task A1x1RangeHoldingAnError_CountsAsASingleCell_ForTheLookupValueRule(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Folded finding I1: VLOOKUP now shares the single-cell lookup-value rule. Before this fix MySheet
    // returned the 1x1 range's collapse artifact #VALUE!; Aspose.Cells 26.7.0 PLAIN / CSE both return the
    // referenced cell's #DIV/0!. Multi-cell collapse artifacts remain pinned above.
    [Test]
    [Arguments("=VLOOKUP(A1:A1,B1:C3,2,FALSE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(A1:A1,B1:D1,1,FALSE)", "#DIV/0!")]
    public async Task TableLookupsOverA1x1RangeHoldingAnError_PropagateTheCellError(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Claim measured FALSE, registered — not fixed here. The brief's I4 wording named ErrCell:ErrCell
    // (a NAME on both sides of ':') as an equivalent 1x1-range shape to A1:A1; it is not. A1:A1 and a name
    // BOUND TO a 1x1 range (ErrRange above) parse to a RangeReference the widened IsLookupValueError now
    // reads; ErrCell:ErrCell parses to an OpenRangeReference with a GARBAGE ColMin/ColMax (measured: both
    // 1766725648 on this fixture) — the parser treats "Identifier:Identifier" as Excel's whole-COLUMN syntax
    // (A:A) generalized to any token before/after the ':', and a non-column-letter name degrades into a
    // meaningless numeric column instead of failing to parse. This is a SEPARATE, pre-existing PARSER defect
    // (measured with a throwaway probe: ExpressionParser.Parse("=ErrCell:ErrCell",...) is an
    // OpenRangeReference, not a DynamicRange over two NameReference endpoints as the finding assumed), well
    // outside "one arm on IsLookupValueError" — registering it for the controller rather than widening this
    // fix wave into the parser. Unaffected by this fix wave: branch stays #N/A; oracle 26.7.0 = 26.6.0: #DIV/0!.
    [Test]
    [Arguments("=MATCH(ErrCell:ErrCell,B1:B3,0)", "#N/A")]
    public async Task NameColonName_ParsesAsAGarbageOpenColumnRange_ARegisteredParserDefect(
        string formula,
        string expected
    )
    {
        await Assert.That(InCell(formula)).IsEqualTo(expected);
    }

    // Direct error table A1, Aspose.Cells 26.7.0 PLAIN / CSE: FALSE mode is #N/A for indexes 0, 1 and 2;
    // TRUE mode is #DIV/0! for indexes 1 and 2. Before this fix MySheet returned #VALUE!, #N/A, #REF!,
    // #N/A and #REF! respectively. Named ErrCell retains its separate CSE-selected #DIV/0! precedence.
    [Test]
    [Arguments("=VLOOKUP(1,A1,0,FALSE)", "#N/A")]
    [Arguments("=VLOOKUP(1,A1,1,FALSE)", "#N/A")]
    [Arguments("=VLOOKUP(1,A1,2,FALSE)", "#N/A")]
    [Arguments("=VLOOKUP(1,A1,1,TRUE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(1,A1,2,TRUE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(1,A1,0,FALSE)", "#N/A")]
    [Arguments("=HLOOKUP(1,A1,1,FALSE)", "#N/A")]
    [Arguments("=HLOOKUP(1,A1,2,FALSE)", "#N/A")]
    [Arguments("=HLOOKUP(1,A1,1,TRUE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(1,A1,2,TRUE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(1,ErrCell,1,FALSE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(1,ErrCell,2,FALSE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(1,ErrCell,1,FALSE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(1,ErrCell,2,FALSE)", "#DIV/0!")]
    public async Task DirectAndNamedErrorTables_KeepTheirMeasuredPrecedence(
        string formula,
        string expected
    ) => await Assert.That(InCell(formula)).IsEqualTo(expected);

    // Fixture addition: A1 is the existing error cell, Z1 = 1/0, and the non-error table cell is B1 = 1.
    // Aspose.Cells 26.7.0 PLAIN / CSE agree. Before this fix, VLOOKUP/HLOOKUP over B1 returned #N/A for
    // error lookup values (and #REF! for index 2) instead of propagating #DIV/0!. The #N/A literal and
    // XLOOKUP twin already agreed with the oracle and guard error identity/shared lookup precedence.
    [Test]
    [Arguments("=VLOOKUP(Z1,B1,1,FALSE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(Z1,B1,1,FALSE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(Z1,B1,1,TRUE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(Z1,B1,1,TRUE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(#N/A,B1,1,FALSE)", "#N/A")]
    [Arguments("=VLOOKUP(ErrLookup,B1,1,FALSE)", "#DIV/0!")]
    [Arguments("=VLOOKUP(Z1,B1,2,FALSE)", "#DIV/0!")]
    [Arguments("=HLOOKUP(Z1,B1,2,FALSE)", "#DIV/0!")]
    [Arguments("=XLOOKUP(Z1,B1,B1)", "#DIV/0!")]
    public async Task ErrorLookupValueOverANonErrorScalarTable_LeadsTheLookup(
        string formula,
        string expected
    ) => await Assert.That(InCell(formula)).IsEqualTo(expected);
}
