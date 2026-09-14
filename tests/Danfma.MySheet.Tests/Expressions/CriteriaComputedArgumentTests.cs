using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Phase 11a Rule B (item 22 of the compatibility sweep): <b>a range slot of <c>SUMIF</c>/<c>SUMIFS</c>/
/// <c>COUNTIF</c>/<c>COUNTIFS</c>/<c>AVERAGEIF</c>/<c>AVERAGEIFS</c>/<c>MAXIFS</c>/<c>MINIFS</c> — criteria range
/// and sum/average/max/min range alike — that is not a reference node and that the mini-CSE would stream is
/// <c>#REF!</c>.</b>
/// <para>
/// Oracle: <b>Aspose.Cells 26.6.0</b>, measured 2026-09-10 on this exact fixture, in BOTH entry modes. For a
/// COMPUTED array in a range slot (<c>A1:A3*1</c>, <c>ROW(A1:A3)</c>, <c>LEN(A1:A3)</c>, …) the plain column is
/// <c>#VALUE!</c> and the array-entered column (<c>Cell.SetArrayFormula(f, 1, 1)</c>) is <c>#REF!</c>; for a
/// PRODUCER in that slot (<c>COUNTIF(FILTER(…),…)</c>, <c>COUNTIF(SEQUENCE(5),">3")</c>,
/// <c>SUMIF(SORT(…),…)</c>) it is <c>#REF!</c> in BOTH columns. The mini-CSE implements the array-entered
/// rule, so <c>#REF!</c> is the target for the computed rows, and it is also the only value the two producer
/// columns agree on — which is why the rule is <c>#REF!</c> and not <c>#VALUE!</c>. The producer rows
/// themselves become pinnable only when Phase 7 ships those functions: Phase 7 item 17 must add them here.
/// </para>
/// <para>
/// Two halves. The <b>acceptance</b> pins are RED at <c>0b93d66</c>: <c>PositionalRange.Open</c> is deliberately
/// not <c>OpenArrayOrRange</c>, so a computed array falls to <c>ArgumentFlattening.ExpandComputedValues</c> and
/// collapses to ONE <c>#VALUE!</c> element; a single-range form then reports an empty scan (a silent 0, or
/// <c>#DIV/0!</c> for AVERAGEIF) and a paired form raises the length mismatch (<c>#VALUE!</c>). The
/// <b>must-not-move</b> pins are GREEN today and must stay green: they prove the gate is exactly
/// <c>ArrayEvaluation.TryStream</c>'s predicate (a non-reference node the mini-CSE would stream) and nothing
/// wider — a scalar-conditioned <c>IF</c>, <c>CHOOSE</c>, <c>OFFSET</c>, a name, a cell, a bare scalar and a
/// refused open range all keep their current answers. The refused open range is an ARRAY that still reaches
/// a scan collapsed, so it is the standing limit of the rule rather than decoration: see
/// <c>PositionalRange.Open</c>'s doc. Either half of a <c>LET</c> was the other standing limit until Phase 11c
/// (array bindings) flipped <see cref="LetBoundComputedArray_InARangeSlot_IsRefused"/> from 0 to <c>#REF!</c>.
/// </para>
/// </summary>
public class CriteriaComputedArgumentTests
{
    // Rule A's fixture, verbatim (DefinedNameArrayEligibilityTests owns its rationale): A1:A3 = 5, 0, 9;
    // B1:B3 = 1, 2, 3; D1:D3 = 1, 22, 333; F1:F2 = 1, 2 and F3 = SUBTOTAL(9,F1:F2). Names are workbook-level
    // and sheet-qualified; GhostName points at a sheet that does not exist and MyCell is a single cell.
    private static (Workbook Workbook, Sheet Sheet) Grid()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);
        sheet["B1"] = new NumberValue(1);
        sheet["B2"] = new NumberValue(2);
        sheet["B3"] = new NumberValue(3);

        sheet["D1"] = new NumberValue(1);
        sheet["D2"] = new NumberValue(22);
        sheet["D3"] = new NumberValue(333);

        sheet["F1"] = new NumberValue(1);
        sheet["F2"] = new NumberValue(2);
        sheet["F3"] = ExpressionParser.Parse("=SUBTOTAL(9,F1:F2)", sheet);

        // The cell-error control for the sweep 34(a) arm: E2 holds a #DIV/0! so the range-slot rows can
        // pin the cell-vs-argument error boundary (an error CELL is content, not the slot's own error).
        sheet["E2"] = ExpressionParser.Parse("=1/0", sheet);

        workbook.DefineName("Rng", "Sheet1!$A$1:$A$3");
        workbook.DefineName("Wide", "Sheet1!$A$1:$B$3");
        workbook.DefineName("MyName", "Sheet1!$D$1:$D$3");
        workbook.DefineName("Nested", "Sheet1!$F$1:$F$3");
        workbook.DefineName("MyCell", "Sheet1!$A$2");
        workbook.DefineName("MyCol", "Sheet1!$A:$A");
        workbook.DefineName("GhostName", "Ghost!$A$1:$A$3");
        workbook.DefineName("UnN", "(Sheet1!$A$1:$A$2,Sheet1!$A$3:$A$3)");
        workbook.DefineName("Threshold", new NumberValue(4));

        return (workbook, sheet);
    }

    private static object? OnGrid(string formula)
    {
        var (workbook, sheet) = Grid();
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static object? OnTextGrid(string formula)
    {
        var (workbook, sheet) = Grid();
        sheet["A2"] = new Danfma.MySheet.Expressions.StringValue("x");
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    private static double Num(object? value) => value is double d ? d : double.NaN;

    // ================================================================================================
    // ACCEPTANCE pins (RED at 0b93d66) — one assertion each, so every row reports its own value.
    // Each comment names the value measured on this tree at 0b93d66 in brackets; the SILENT rows (a
    // plausible NUMBER, not an error) are the reason the item exists, because a shipped FILTER behind them
    // would be silently wrong. Aspose 26.6.0: #VALUE! plain / #REF! array-entered unless a comment says
    // otherwise.
    // ================================================================================================

    // ================================================================================================
    // Sweep item 34(a) — an error-valued argument in a range slot PROPAGATES its own error: the ONE
    // general arm in PositionalRange.Open's fallback (RangeValueCursor's mirror for COUNTIF, the
    // COUNTBLANK slot included). Oracle Aspose.Cells 26.6.0, measured 2026-09-11, PLAIN and
    // array-entered agreeing on every row. What each row WAS: a silent 0 (the error streamed as the one
    // element the criteria discards), a wrong #DIV/0! where AVERAGEIF divided its empty scan, or the
    // paired form's #VALUE! length mismatch — every bracketed number below.
    // ================================================================================================

    [Test]
    public async Task AnErrorValuedArgument_InARangeSlot_PropagatesItsError()
    {
        // The two rows the sweep recorded: #NAME? [was 0] and #DIV/0! [was 0].
        await Assert.That(OnGrid("=COUNTIF(NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=COUNTIF(1/0,\">0\")")).IsEqualTo(ErrorValue.DivByZero);

        // SUMIF's range slot: #NAME? [was 0] / #DIV/0! [was 0].
        await Assert.That(OnGrid("=SUMIF(NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=SUMIF(1/0,\">0\")")).IsEqualTo(ErrorValue.DivByZero);

        // AVERAGEIF: #NAME? [was #DIV/0! — the empty scan of a single-range AVERAGEIF].
        await Assert.That(OnGrid("=AVERAGEIF(NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);

        // COUNTIFS: #NAME? [was 0] / #DIV/0! [was 0].
        await Assert.That(OnGrid("=COUNTIFS(NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=COUNTIFS(1/0,\">0\")")).IsEqualTo(ErrorValue.DivByZero);

        // The paired forms: the error leads the length validation — #NAME?/#DIV/0! [both were #VALUE!].
        await Assert.That(OnGrid("=SUMIFS(B1:B3,NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=SUMIFS(B1:B3,1/0,\">0\")")).IsEqualTo(ErrorValue.DivByZero);
        await Assert.That(OnGrid("=AVERAGEIFS(B1:B3,NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=MAXIFS(B1:B3,NoSuch,\">0\")")).IsEqualTo(ErrorValue.Name);

        // The VALUE slot: #NAME? [was 0] / #DIV/0! [was 0].
        await Assert.That(OnGrid("=SUMIF(A1:A3,\">0\",NoSuch)")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=SUMIF(A1:A3,\">0\",1/0)")).IsEqualTo(ErrorValue.DivByZero);

        // COUNTBLANK, the family's third member: #NAME? [was 0] / #DIV/0! [was 0].
        await Assert.That(OnGrid("=COUNTBLANK(NoSuch)")).IsEqualTo(ErrorValue.Name);
        await Assert.That(OnGrid("=COUNTBLANK(1/0)")).IsEqualTo(ErrorValue.DivByZero);
    }

    [Test]
    public async Task AScalarBoundName_InARangeSlot_StillScansAsItsValue()
    {
        // The arm's guard rail: only an ERROR-valued argument propagates. Threshold is bound to the
        // constant 4, so its one element matches ">0" and COUNTIF is 1 — unchanged. (The oracle answers
        // #REF! here, the RECORDED scalar-slot divergence this file pins below with the literal 5 — a
        // scalar in a range slot is not this arm's; the arm must not widen it into an error.)
        await Assert.That(Num(OnGrid("=COUNTIF(Threshold,\">0\")"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task AnErrorCell_InARangeSlot_IsContent_NotTheArgumentsOwnError()
    {
        // The other guard rail, the cell-vs-argument boundary (Aspose 26.6.0, 2026-09-11, both entry
        // modes: 0): a reference whose CELL holds a #DIV/0! does not propagate it — the node's own value
        // is the reference, the error belongs to the value walk, and the criteria simply does not match
        // the error cell. The arm's IsOwnSlotError guard is what keeps this row at its 0.
        await Assert.That(Num(OnGrid("=COUNTIF(E2,\">0\")"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task Sumif_OverAComputedCriteriaRange_IsRef()
    {
        // SILENT [0]: the lone #VALUE! element matches no criterion, so the scan is empty.
        await Assert.That(OnGrid("=SUMIF(A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countif_OverAComputedCriteriaRange_IsRef()
    {
        // SILENT [0].
        await Assert.That(OnGrid("=COUNTIF(A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countifs_OverAComputedCriteriaRange_IsRef()
    {
        // SILENT [0]: COUNTIFS' first argument IS the criteria range, so a single pair has no second length
        // to disagree with and it belongs to the silent half despite its plural name.
        await Assert.That(OnGrid("=COUNTIFS(A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task AverageIf_OverAComputedCriteriaRange_IsRef()
    {
        // [#DIV/0!]: an empty scan divided by a zero count — an error, but the WRONG one.
        await Assert.That(OnGrid("=AVERAGEIF(A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Sumifs_OverAComputedCriteriaRange_IsRef()
    {
        // [#VALUE!]: the paired form sees 1 element against 3 and raises the length mismatch.
        await Assert.That(OnGrid("=SUMIFS(B1:B3,A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task AverageIfs_OverAComputedCriteriaRange_IsRef()
    {
        // [#VALUE!] — the length mismatch.
        await Assert
            .That(OnGrid("=AVERAGEIFS(B1:B3,A1:A3*1,\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task MaxIfs_OverAComputedCriteriaRange_IsRef()
    {
        // [#VALUE!] — the length mismatch.
        await Assert.That(OnGrid("=MAXIFS(B1:B3,A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task MinIfs_OverAComputedCriteriaRange_IsRef()
    {
        // [#VALUE!] — the length mismatch.
        await Assert.That(OnGrid("=MINIFS(B1:B3,A1:A3*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Sumif_OverAComputedSumRange_IsRef()
    {
        // SILENT [0]: the VALUE slot, not the criteria slot — the criteria range is real, two cells match,
        // and the collapsed sum range contributes nothing numeric at their positions.
        await Assert.That(OnGrid("=SUMIF(A1:A3,\">0\",B1:B3*1)")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task AverageIf_OverAComputedAverageRange_IsRef()
    {
        // [#DIV/0!]: the VALUE slot of AVERAGEIF.
        await Assert
            .That(OnGrid("=AVERAGEIF(A1:A3,\">0\",B1:B3*1)"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Sumifs_OverAComputedValueRange_IsRef()
    {
        // [#VALUE!]: the VALUE slot of SUMIFS (its FIRST argument).
        await Assert.That(OnGrid("=SUMIFS(A1:A3*1,B1:B3,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task MaxIfs_OverAComputedValueRange_IsRef()
    {
        // [#VALUE!]: the VALUE slot of MAXIFS.
        await Assert.That(OnGrid("=MAXIFS(B1:B3*1,A1:A3,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countifs_OverAComputedSecondCriteriaRange_IsRef()
    {
        // [#VALUE!]: the computed array is the SECOND pair's range, so the gate must check every pair, not
        // only the first.
        await Assert
            .That(OnGrid("=COUNTIFS(A1:A3,\">0\",B1:B3*1,\">1\")"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countif_OverARowVector_IsRef()
    {
        // SILENT [0]. Aspose 26.6.0: #REF! in BOTH modes — the one computed row whose two columns agree.
        await Assert.That(OnGrid("=COUNTIF(ROW(A1:A3),\">1\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countif_OverANegatedRange_IsRef()
    {
        // SILENT [0]: a lifted unary is array-eligible exactly like a binary operation.
        await Assert.That(OnGrid("=COUNTIF(-A1:A3,\"<0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countif_OverALiftedFunction_IsRef()
    {
        // SILENT [0]: a lifted pure-scalar built-in (LEN) is array-eligible too.
        await Assert.That(OnGrid("=COUNTIF(LEN(A1:A3),\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Countif_OverANameBuiltArray_IsRef()
    {
        // SILENT [1]: at 0b93d66 the name is an opaque truthy reference value, so (Rng<>0)*1 is ONE element
        // equal to 1 and COUNTIF counts it. Rule A first makes the operand array-eligible; Rule B then rejects
        // it — this row needs both.
        await Assert.That(OnGrid("=COUNTIF((Rng<>0)*1,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task CriteriaFamily_AcceptsAnArrayConditionedReferenceSelector()
    {
        // A1:A3=5,0,9 and B1:B3=1,2,3; Aspose.Cells 26.7.0 CSE. Every row was #REF! -> its listed CSE
        // value after the criteria gate began selecting the reference branch from the first condition element.
        await Assert.That(Num(OnGrid("=COUNTIF(IF(A1:A3>0,A1:A3),\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUMIF(IF(A1:A3>0,A1:A3),\">0\")"))).IsEqualTo(14.0);
        await Assert.That(Num(OnGrid("=AVERAGEIF(IF(A1:A3>0,A1:A3),\">0\")"))).IsEqualTo(7.0);
        await Assert.That(Num(OnGrid("=COUNTIFS(IF(A1:A3>0,A1:A3),\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUMIFS(B1:B3,IF(A1:A3>0,A1:A3),\">0\")"))).IsEqualTo(4.0);
        await Assert
            .That(Num(OnGrid("=AVERAGEIFS(B1:B3,IF(A1:A3>0,A1:A3),\">0\")")))
            .IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=COUNTIF(IF(A1:A3>4,A1:A3,B1:B3),\">0\")"))).IsEqualTo(2.0);
    }

    [Test]
    [Arguments("=COUNTIF(IF(A1:A3>0,A1:A3),\">0\")", 2.0)]
    [Arguments("=SUMIF(IF(A1:A3>0,A1:A3),\">0\")", 14.0)]
    [Arguments("=COUNTIF(IF(A1:A3>4,A1:A3,B1:B3),\">0\")", 2.0)]
    [Arguments("=COUNTIFS(IF(A1:A3>0,A1:A3),\">0\")", 2.0)]
    [Arguments("=SUMIFS(B1:B3,IF(A1:A3>0,A1:A3),\">0\")", 4.0)]
    [Arguments("=AVERAGEIF(IF(A1:A3>0,A1:A3),\">0\")", 7.0)]
    [Arguments("=AVERAGEIFS(B1:B3,IF(A1:A3>0,A1:A3),\">0\")", 2.0)]
    public async Task CriteriaFamily_ArrayConditionedSelector_TextCellFixture(
        string formula,
        double expected
    )
    {
        // A1:A3=5,"x",9 and B1:B3=1,2,3; Aspose.Cells 26.7.0 CSE. Each row was #REF! -> expected:
        // text compares greater than zero, so the selected A reference retains all three positions.
        await Assert.That(Num(OnTextGrid(formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task CriteriaFamily_ArrayConditionedSelector_CoversEveryRangeSlot()
    {
        // A1:A3=5,0,9 and B1:B3=1,2,3; Aspose.Cells 26.7.0 CSE. These were #REF! -> 3, 1, 0, 4:
        // the selected reference is accepted in max/min, COUNTBLANK, and the SUMIFS value-range slot.
        await Assert
            .That(Num(OnGrid("=MAXIFS(IF(A1:A3>0,B1:B3,A1:A3),A1:A3,\">0\")")))
            .IsEqualTo(3.0);
        await Assert
            .That(Num(OnGrid("=MINIFS(IF(A1:A3>0,B1:B3,A1:A3),A1:A3,\">0\")")))
            .IsEqualTo(1.0);
        await Assert.That(Num(OnGrid("=COUNTBLANK(IF(A1:A3>0,A1:A3,B1:B3))"))).IsEqualTo(0.0);
        await Assert
            .That(Num(OnGrid("=SUMIFS(IF(A1:A3>0,B1:B3,A1:A3),A1:A3,\">0\")")))
            .IsEqualTo(4.0);
    }

    [Test]
    public async Task CriteriaFamily_ArrayConditionedSelector_EvaluatesAVolatileConditionOnce()
    {
        // 200 seeded evaluations. A1:A3=5,0,9 / B1:B3=1,2,3; a single TICK selects A (2 matches) or B
        // (3 matches). A second condition evaluation would increment draws past one and is rejected directly.
        for (var seed = 0; seed < 200; seed++)
        {
            var (workbook, sheet) = Grid();
            var draws = 0;
            var first = seed % 2 == 0;
            workbook.RegisterFunction(
                "TICK",
                (_, _) =>
                {
                    draws++;
                    return first ? 1 : 0;
                }
            );

            var result = ExpressionParser
                .Parse("=COUNTIF(IF(TICK()>0,A1:A3,B1:B3),\">0\")", sheet)
                .Evaluate(workbook)
                .AsObject();

            await Assert.That(Num(result)).IsEqualTo(first ? 2.0 : 3.0);
            await Assert.That(draws).IsEqualTo(1);
        }
    }

    [Test]
    public async Task CriteriaFamily_AcceptsAMixedSelectorOnlyWhenItChoosesAReference()
    {
        // Measured on Aspose.Cells 26.7.0, one formula per workbook, PLAIN/CSE respectively:
        // TRUE/ref/producer and FALSE/producer/ref are 2/2 (COUNTIF) and 0/0 (COUNTBLANK). Reversing the
        // selected branch gives #REF!/#REF!. MySheet previously returned #REF! for all eight shapes.
        await Assert
            .That(Num(OnGrid("=COUNTIF(IF(TRUE,A1:A3,SEQUENCE(3)),\">0\")")))
            .IsEqualTo(2.0);
        await Assert
            .That(Num(OnGrid("=COUNTIF(IF(FALSE,SEQUENCE(3),A1:A3),\">0\")")))
            .IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=COUNTBLANK(IF(TRUE,A1:A3,SEQUENCE(3)))"))).IsEqualTo(0.0);
        await Assert.That(Num(OnGrid("=COUNTBLANK(IF(FALSE,SEQUENCE(3),A1:A3))"))).IsEqualTo(0.0);
        await Assert
            .That(OnGrid("=COUNTIF(IF(FALSE,A1:A3,SEQUENCE(3)),\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnGrid("=COUNTIF(IF(TRUE,SEQUENCE(3),A1:A3),\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
    }

    // ================================================================================================
    // MUST-NOT-MOVE pins — GREEN today and after. The gate is exactly TryStream's predicate, so every
    // shape below is outside it: not array-eligible (a scalar-conditioned IF, CHOOSE, OFFSET, a bare
    // scalar), a reference node (a name, a cell, an open range), or refused by the cost guard.
    // ================================================================================================

    [Test]
    public async Task ReferenceReturningShapes_InARangeSlot_AreUntouched()
    {
        // Aspose 26.6.0: 2 / 2 on all three rows (plain / array-entered). The scalar-conditioned IF row
        // arrived here pinned at MySheet's 0 and CLOSED as sweep item 32 (was 0, now the oracle's 2, both
        // modes): an all-bare-reference-branch IF is a reference at this top level — the gate excludes it
        // and the fallback reads the taken branch's range. CHOOSE and OFFSET already returned references
        // and are read as ranges.
        await Assert.That(Num(OnGrid("=COUNTIF(IF(TRUE,A1:A3,B1:B3),\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=COUNTIF(CHOOSE(1,A1:A3,B1:B3),\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=COUNTIF(OFFSET(A1,0,0,3,1),\">0\")"))).IsEqualTo(2.0);
    }

    [Test]
    public async Task NamesAndCells_InARangeSlot_AreUntouched()
    {
        // A bare name is a reference node and stays on the reference path (Rule A's top-level rule), a cell
        // and an open range are reference nodes outright, and a name on a missing sheet is the structural
        // #REF! the family already reports. Aspose 26.6.0, plain / array-entered: 2 / 2, 4 / 4, 1 / 1,
        // 3 / 3, 3 / 3, 1 / 1, 2 / 2, 0 / 0, #REF! / #REF!.
        await Assert.That(Num(OnGrid("=COUNTIF(Rng,\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUMIF(Rng,\">0\",B1:B3)"))).IsEqualTo(4.0);
        await Assert.That(Num(OnGrid("=MINIFS(B1:B3,Rng,\">0\")"))).IsEqualTo(1.0);
        await Assert.That(Num(OnGrid("=MAXIFS(B1:B3,A1:A3,\">0\")"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=SUMIFS(B1:B3,A1:A3,\">0\",B1:B3,\">1\")"))).IsEqualTo(3.0);
        await Assert.That(Num(OnGrid("=COUNTIF(A1,\">0\")"))).IsEqualTo(1.0);
        await Assert.That(Num(OnGrid("=COUNTIF(A:A,\">0\")"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=COUNTIF(MyCell,\">0\")"))).IsEqualTo(0.0);
        await Assert.That(OnGrid("=COUNTIF(GhostName,\">0\")")).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task Sumproduct_KeepsItsOptInArrayFactory()
    {
        // SUMPRODUCT is the one member of the positional-scan family that opted IN to computed arrays
        // (PositionalRange.OpenArrayOrRange); Rule B gates the eight Open calls, not that factory.
        // Aspose 26.6.0: 2 / 2 and 32 / 32.
        await Assert.That(Num(OnGrid("=SUMPRODUCT((A1:A3<>0)*1)"))).IsEqualTo(2.0);
        await Assert.That(Num(OnGrid("=SUMPRODUCT(A1:A3*1,B1:B3)"))).IsEqualTo(32.0);
    }

    [Test]
    public async Task ScalarShapes_InARangeSlot_AreUnchanged_KnownDivergence()
    {
        // DELIBERATE DIVERGENCE, left as it is by this phase: a bare SCALAR in a range slot is the sweep's
        // scalar rule, not this item's, and no array producer ever takes that shape. A1*1 and 5 are not
        // array-eligible, so the gate does not see them. Aspose 26.6.0: #REF! PLAIN and #REF! ARRAY-ENTERED
        // for both (the two columns agree; MySheet answers 1 — the scalar 5 matches ">0").
        await Assert.That(Num(OnGrid("=COUNTIF(A1*1,\">0\")"))).IsEqualTo(1.0);
        await Assert.That(Num(OnGrid("=COUNTIF(5,\">0\")"))).IsEqualTo(1.0);
    }

    [Test]
    public async Task OpenRangeArithmetic_InARangeSlot_IsUnchanged_KnownDivergence()
    {
        // DELIBERATE DIVERGENCE, left as it is by this phase: the mini-CSE's cost guard REFUSES a
        // whole-column operand, so A:A*1 is not "array-eligible" and the gate does not see it; the argument
        // collapses to one #VALUE! element and the scan is empty. Aspose 26.6.0: #REF! PLAIN and #REF!
        // ARRAY-ENTERED (the two columns agree; MySheet answers 0).
        await Assert.That(Num(OnGrid("=SUMIF(A:A*1,\">0\")"))).IsEqualTo(0.0);
    }

    [Test]
    public async Task LetBoundComputedArray_InARangeSlot_IsRefused()
    {
        // Aspose 26.6.0, measured 2026-09-10 and re-measured 2026-09-11 (ArrayBindingTests' probe), plain /
        // array-entered: #VALUE! / #REF! for all four rows — the same split every COMPUTED array in a range
        // slot shows (this file's header), so the CSE column, #REF!, is the target.
        //
        // Pinned at 0 by Phase 11a as a PRE-EXISTING DIVERGENCE (measured identical on main at 5f9d1ac and
        // on that branch), flipped to #REF! by Phase 11c (array bindings): 0 → #REF! for every row. What made
        // it 0: Rule B's gate was `!IsBareReferenceNode(argument) && IsArrayEligible(argument, context)`
        // (context-free then), and
        // a LET escaped it from BOTH sides. LET(...) written in the slot was a Let NODE the shape probe
        // treated as an opaque scalar, so it was not array-eligible; a LET-BOUND NAME in the slot was a
        // reference node IsBareReferenceNode admitted by design (Rule A's top-level rule), and its binding
        // had already been collapsed to one scalar by NamedReferences.CaptureValue, which evaluated a
        // non-range binding rather than capturing a reference. Either way the argument reached
        // PositionalRange.Open as a single #VALUE! element and the scan was empty — the same silent 0 the
        // gate removed everywhere it could see. Phase 11c makes a Let node array-eligible and makes the
        // bare-reference predicate context-aware (a name bound to an array is not a reference), so the gate
        // sees both shapes and refuses them. The producer twin of these rows —
        // LET(f,FILTER(...),COUNTIF(f,...)) — is DynamicArrayTests.ALetBoundProducer_StreamsWhole_… and
        // ArrayBindingTests.
        await Assert
            .That(OnGrid("=COUNTIF(LET(r,A1:A3,r*1),\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnGrid("=LET(r,A1:A3*1,COUNTIF(r,\">0\"))"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnGrid("=SUMIF(LET(r,A1:A3,r*1),\">0\")"))
            .IsEqualTo(ErrorValue.Reference);
        await Assert
            .That(OnGrid("=LET(r,A1:A3*1,SUMIF(r,\">0\"))"))
            .IsEqualTo(ErrorValue.Reference);
    }
}
