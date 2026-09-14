using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Sweep item 37 (Bug 6): <c>INDEX(area, 0, n)</c> / <c>INDEX(area, n, 0)</c> / <c>INDEX(area, 0, 0)</c> /
/// the 2-arg <c>INDEX(area, 0)</c> over a one-column or one-row area return the area's column / row / whole
/// extent as a REFERENCE, not <c>#REF!</c> — so every reference-aware consumer (SUM, COUNT, COUNTA,
/// SUMPRODUCT, AGGREGATE, the COUNTIF family, ROWS, COLUMNS, ROW, COLUMN, AREAS, ISREF, MATCH's lookup
/// array, a nested INDEX, an OFFSET base, a bare cell's implicit intersection, a ':' range endpoint) sees a
/// real range, not <c>#REF!</c> and not a materialized value.
///
/// <para>Oracle: Aspose.Cells 26.7.0, one formula per workbook, PLAIN and CSE agreeing everywhere measured
/// (2026-09-14; CSE spot-checked on SUM/ROWS/the row-form SUM/COUNT/MATCH/both-zero SUM — all identical to
/// PLAIN, except the bare-cell CSE reading, which takes the array's TOP-LEFT rather than intersecting by
/// the formula cell's position — a pre-existing PLAIN-vs-CSE distinction this item does not need to model:
/// the divergence probe and this suite both read cells PLAIN).</para>
///
/// <para>The idx fixture mirrors the corpus's own shape: <c>Main!E5:H10</c>, every cell in row <c>r</c>
/// equal to <c>r</c> (rows 5..10) across columns E-H — so column E (E5:E10) is <c>[5,6,7,8,9,10]</c> and row
/// 5 (E5:H5) is <c>[5,5,5,5]</c>.</para>
/// </summary>
public class IndexZeroAxisTests
{
    private static Workbook IdxFixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");

        foreach (var row in Enumerable.Range(5, 6))
        foreach (var column in new[] { "E", "F", "G", "H" })
        {
            sheet[$"{column}{row}"] = new NumberValue(row);
        }

        return workbook;
    }

    private static Workbook BasicFixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["A1"] = new NumberValue(5);
        sheet["A2"] = new NumberValue(0);
        sheet["A3"] = new NumberValue(9);

        return workbook;
    }

    private static Workbook OneRowFixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["B1"] = new NumberValue(1);
        sheet["C1"] = new NumberValue(2);
        sheet["D1"] = new NumberValue(3);

        return workbook;
    }

    private static Workbook CriteriaFixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        sheet["A4"] = new NumberValue(4);
        sheet["B1"] = new NumberValue(1);
        sheet["B2"] = new NumberValue(1);
        sheet["B3"] = new NumberValue(1);
        sheet["B4"] = new NumberValue(1);

        return workbook;
    }

    // Stores the formula at the given cell (default far from every fixture, matching the project's
    // "AZ5000"-style out-of-the-way convention) and reads it back through the real cell boundary (implicit
    // intersection included).
    private static object? Eval(
        Workbook workbook,
        string formula,
        string sheet = "Main",
        string cell = "Z100"
    )
    {
        var target = workbook[sheet];
        target[cell] = ExpressionParser.Parse(formula, target);

        return workbook.GetCellValue(sheet, cell).AsObject();
    }

    private static double? Num(object? value) => value as double?;

    // === Column form: INDEX(E5:H10,0,1) = column E, E5:E10 = [5,6,7,8,9,10] ===============================

    [Test]
    [Arguments("=SUM(INDEX(E5:H10,0,1))", 45.0)]
    [Arguments("=COUNT(INDEX(E5:H10,0,1))", 6.0)]
    [Arguments("=COUNTA(INDEX(E5:H10,0,1))", 6.0)]
    [Arguments("=SUMPRODUCT(INDEX(E5:H10,0,1))", 45.0)]
    [Arguments("=AGGREGATE(9,6,INDEX(E5:H10,0,1))", 45.0)]
    [Arguments("=AGGREGATE(15,6,INDEX(E5:H10,0,1),1)", 5.0)]
    [Arguments("=COUNTIF(INDEX(E5:H10,0,1),\">6\")", 4.0)]
    [Arguments("=SUMIF(INDEX(E5:H10,0,1),\">6\")", 34.0)]
    [Arguments("=COUNTIFS(INDEX(E5:H10,0,1),\">6\")", 4.0)]
    [Arguments("=ROWS(INDEX(E5:H10,0,1))", 6.0)]
    [Arguments("=COLUMNS(INDEX(E5:H10,0,1))", 1.0)]
    [Arguments("=ROW(INDEX(E5:H10,0,1))", 5.0)]
    [Arguments("=COLUMN(INDEX(E5:H10,0,1))", 5.0)]
    [Arguments("=AREAS(INDEX(E5:H10,0,1))", 1.0)]
    [Arguments("=MATCH(7,INDEX(E5:H10,0,1),0)", 3.0)]
    [Arguments("=INDEX(INDEX(E5:H10,0,1),2,1)", 6.0)] // nested INDEX
    [Arguments("=OFFSET(INDEX(E5:H10,0,1),1,0)", 6.0)] // OFFSET base: E5 + (1,0) = E6
    [Arguments("=ROW(INDEX(E5:H10,0,1))-4", 1.0)] // lifted under a scalar operator
    [Arguments("=SUM(INDEX(E5:H10,0,1):H10)", 180.0)] // ':' range endpoint: bounding box E5:H10
    public async Task ColumnForm_NumericConsumers_MatchTheOracle(string formula, double expected)
    {
        await Assert.That(Num(Eval(IdxFixture(), formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task ColumnForm_IsRef_IsTrue()
    {
        await Assert.That(Eval(IdxFixture(), "=ISREF(INDEX(E5:H10,0,1))") as bool?).IsTrue();
    }

    [Test]
    public async Task ColumnForm_BareCell_ImplicitlyIntersectsByTheFormulaCellsRow()
    {
        var workbook = IdxFixture();

        // Row 7 IS inside E5:H10's rows (5..10): intersect -> that row's cell in column E -> E7 -> 7.
        await Assert.That(Num(Eval(workbook, "=INDEX(E5:H10,0,1)", cell: "J7"))).IsEqualTo(7.0);

        // Row 60 is outside 5..10: the single-column reference misses the formula's line -> #VALUE!.
        await Assert
            .That(Eval(workbook, "=INDEX(E5:H10,0,1)", cell: "H60"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // === Row form: INDEX(E5:H10,1,0) = row 5, E5:H5 = [5,5,5,5] ============================================

    [Test]
    [Arguments("=SUM(INDEX(E5:H10,1,0))", 20.0)]
    [Arguments("=COUNT(INDEX(E5:H10,1,0))", 4.0)]
    [Arguments("=COUNTA(INDEX(E5:H10,1,0))", 4.0)]
    [Arguments("=AGGREGATE(9,6,INDEX(E5:H10,1,0))", 20.0)]
    [Arguments("=ROWS(INDEX(E5:H10,1,0))", 1.0)]
    [Arguments("=COLUMNS(INDEX(E5:H10,1,0))", 4.0)]
    [Arguments("=ROW(INDEX(E5:H10,1,0))", 5.0)]
    [Arguments("=COLUMN(INDEX(E5:H10,1,0))", 5.0)]
    public async Task RowForm_NumericConsumers_MatchTheOracle(string formula, double expected)
    {
        await Assert.That(Num(Eval(IdxFixture(), formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task RowForm_BareCell_ImplicitlyIntersectsByTheFormulaCellsColumn()
    {
        var workbook = IdxFixture();

        // Column F IS inside E5:H10's columns (E..H): intersect -> that column's cell in row 5 -> F5 -> 5.
        await Assert.That(Num(Eval(workbook, "=INDEX(E5:H10,1,0)", cell: "F60"))).IsEqualTo(5.0);

        // Column Z is outside E..H: the single-row reference misses the formula's line -> #VALUE!.
        await Assert
            .That(Eval(workbook, "=INDEX(E5:H10,1,0)", cell: "Z20"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // === Both zero: INDEX(area,0,0) is the WHOLE area =======================================================

    [Test]
    [Arguments("=SUM(INDEX(E5:H10,0,0))", 180.0)]
    [Arguments("=ROWS(INDEX(E5:H10,0,0))", 6.0)]
    [Arguments("=COLUMNS(INDEX(E5:H10,0,0))", 4.0)]
    [Arguments("=COUNT(INDEX(E5:H10,0,0))", 24.0)]
    public async Task BothZero_IsTheWholeArea(string formula, double expected)
    {
        await Assert.That(Num(Eval(IdxFixture(), formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task BothZero_BareCell_Is2DSoThereIsNoSingleIntersection()
    {
        // A rectangle wider than one cell on BOTH axes has no single answer, wherever the formula sits.
        await Assert
            .That(Eval(IdxFixture(), "=INDEX(E5:H10,0,0)", cell: "H24"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // === The 2-arg zero form: INDEX(area,0) over a ONE-COLUMN / ONE-ROW area ================================

    [Test]
    public async Task TwoArgZero_OverAOneColumnArea_IsTheWholeColumn()
    {
        // A distinct cell per assertion: the same cell re-evaluated twice would read the FIRST formula's
        // cached value instead of re-parsing the second (a test-fixture pitfall, not an engine one — every
        // other helper in this suite creates a fresh workbook or cell per formula for the same reason).
        var workbook = BasicFixture();

        await Assert
            .That(Num(Eval(workbook, "=SUM(INDEX(A1:A3,0))", cell: "Z100")))
            .IsEqualTo(14.0);
        await Assert
            .That(Num(Eval(workbook, "=ROWS(INDEX(A1:A3,0))", cell: "Z101")))
            .IsEqualTo(3.0);
        // The formula cell (Z100, outside rows 1..3) has no intersection.
        await Assert
            .That(Eval(workbook, "=INDEX(A1:A3,0)", cell: "Z102"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task TwoArgZero_OverAOneRowArea_IsTheWholeRow()
    {
        var workbook = OneRowFixture();

        await Assert.That(Num(Eval(workbook, "=SUM(INDEX(B1:D1,0))", cell: "Z100"))).IsEqualTo(6.0);
        await Assert
            .That(Num(Eval(workbook, "=COLUMNS(INDEX(B1:D1,0))", cell: "Z101")))
            .IsEqualTo(3.0);
        // The formula cell (Z100, outside columns B..D) has no intersection.
        await Assert
            .That(Eval(workbook, "=INDEX(B1:D1,0)", cell: "Z102"))
            .IsEqualTo(ErrorValue.NotValue);
    }

    // === The COUNTIF-family range slot, over a computed criteria fixture (Bug 9's shape) ====================

    [Test]
    [Arguments("=COUNTIF(INDEX(A1:B4,0,1),\">2\")", 2.0)]
    [Arguments("=SUMIF(INDEX(A1:B4,0,1),\">2\")", 7.0)]
    [Arguments("=COUNTIFS(INDEX(A1:B4,0,1),\">2\")", 2.0)]
    public async Task CriteriaRangeSlot_MatchesTheOracle(string formula, double expected)
    {
        await Assert.That(Num(Eval(CriteriaFixture(), formula))).IsEqualTo(expected);
    }

    // === The corpus AGGREGATE shape (the plan's Phase 2 acceptance row; divergence-probe row 44) ============

    [Test]
    public async Task CorpusAggregateShape_MatchesTheOracle()
    {
        var workbook = IdxFixture();

        await Assert
            .That(
                Num(
                    Eval(
                        workbook,
                        "=IFERROR(AGGREGATE(15,6,(ROW(INDEX(E5:H10,0,MATCH(7,E7:H7,0)))-4)/(INDEX(E5:H10,0,1)<>\"\"),1),\"\")"
                    )
                )
            )
            .IsEqualTo(1.0);
    }

    // === Registered gaps: the zero-axis fix does not close these — a DIFFERENT, pre-existing engine
    // convention decides them, exactly like EmptyTableReferenceTests' "Rows that depend on an ordinary-range
    // gap". Both numbers, oracle Aspose.Cells 26.7.0, measured 2026-09-14, one formula per workbook, PLAIN.
    //
    //   • SUM(INDEX($5:$10,0,5)) — a whole-ROW base. INDEX resolves an open range to its POPULATED bounding
    //     box (NamedReferences.TryResolveReference's boundOpenRanges:true, used by every INDEX/VLOOKUP/OFFSET
    //     base — see WholeColumnConsumerTests.Index_WholeColumn_ByPopulatedPosition, a DELIBERATE, already-
    //     shipped convention this item does not touch). Only E..H are populated in row 5..10, so "column 5"
    //     is the 5th column of that 4-column box — out of bounds — #REF!, where Excel's ungridded column 5
    //     is the absolute column E (45). Not sweep item 37's gap: the SAME box-relative counting already
    //     applies to a NON-zero INDEX($5:$10,r,c) (untouched by this fix) and to VLOOKUP/OFFSET over any
    //     open range.
    //   • AGGREGATE(15,6,(ROW(INDEX(E5:H10,0,1))-ROW(INDEX(INDEX(E5:H10,0,1),1,1))+1)/((INDEX(E5:H10,0,1)<>"")
    //     *(INDEX(E5:H10,0,1)>6)),1) — the divergence probe's row 43 (a MYSHEET-CALC-DIVERGENCES.md formula,
    //     not the plan's own corpus shape, which is CorpusAggregateShape_MatchesTheOracle above and DOES
    //     match). Its denominator compares a BARE INDEX(...) result (not wrapped in ROW/COLUMN) against ""
    //     and 6, and needs those comparisons to run ELEMENT-WISE over the 6-cell column even in PLAIN entry
    //     (Aspose: 3, the SMALL of {3,4,5,6} once the two #DIV/0!s are ignored). Making a BARE INDEX result
    //     array-eligible everywhere a comparison/operator could reach it is not a narrow fix like ROW/COLUMN's
    //     own dedicated single-resolution arm (ArrayEvaluation.TryBuildIndexPositionOperand): it would also
    //     change every EXISTING, non-zero INDEX(range,n) used as a bare consumer argument today (e.g.
    //     NumericAggregation.Fold's AddDirect-vs-AddReferenced split for a referenced TEXT cell), a change
    //     with a blast radius well outside "INDEX with row or column 0 returns a reference". Left as MySheet's
    //     own answer (which, entering this fix, already changed from a stable #REF! to 1 — see checkpoint —
    //     since INDEX now evaluates to a scalar-broadcast Reference instead of throwing #REF! at Evaluate's
    //     top). Registered for a future item, not fixed here. The SAME gap surfaces in a second, simpler
    //     shape — SUMPRODUCT(INDEX(E5:H10,0,1)*1), Aspose 45, MySheet #VALUE! — SUMPRODUCT is another
    //     "always array, regardless of entry mode" consumer (like AGGREGATE's array form); the bare
    //     SUMPRODUCT(INDEX(E5:H10,0,1)) with no operator, listed as a required consumer, is unaffected and
    //     DOES match (ColumnForm_NumericConsumers_MatchTheOracle).
    [Test]
    public async Task WholeRowOpenBase_KeepsTheEnginesOwnAnswer_ADivergence()
    {
        await Assert
            .That(Eval(IdxFixture(), "=SUM(INDEX($5:$10,0,5))"))
            .IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task DivergenceProbeRow43Shape_KeepsTheEnginesOwnAnswer_ADivergence()
    {
        var workbook = IdxFixture();

        await Assert
            .That(
                Num(
                    Eval(
                        workbook,
                        "=AGGREGATE(15,6,(ROW(INDEX(E5:H10,0,1))-ROW(INDEX(INDEX(E5:H10,0,1),1,1))+1)/((INDEX(E5:H10,0,1)<>\"\")*(INDEX(E5:H10,0,1)>6)),1)"
                    )
                )
            )
            .IsEqualTo(1.0);
    }
}
