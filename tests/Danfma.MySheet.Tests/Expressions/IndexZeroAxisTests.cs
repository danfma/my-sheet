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
/// <para>Oracle: Aspose.Cells 26.7.0, one formula per workbook (2026-09-14). Where PLAIN and CSE split,
/// MySheet follows CSE: the <c>OFFSET(INDEX(...))</c> base, operators over the zero-axis reference, and
/// bare-cell readings split. The divergence probe and this suite both read bare cells PLAIN.</para>
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

    private static Workbook ThreeByThreeFixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        var value = 1;

        foreach (var row in Enumerable.Range(1, 3))
        foreach (var column in new[] { "A", "B", "C" })
        {
            sheet[$"{column}{row}"] = new NumberValue(value++);
        }

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

    // Aspose.Cells 26.7.0 returns #VALUE! for a negative truncated axis, including a 1x1 table;
    // before this fix the scalar path returned #REF!. Positive indices past the extent remain #REF!.
    [Test]
    [Arguments("=INDEX(A1,-1)", "#VALUE!")]
    [Arguments("=INDEX(A1,1,-1)", "#VALUE!")]
    [Arguments("=INDEX(A1,2)", "#REF!")]
    [Arguments("=INDEX(A1:C3,-1,1)", "#VALUE!")]
    [Arguments("=SUM(INDEX(A1:C3,-0.5,1))", "12")]
    public async Task NegativeAxes_UseTheSameValidationForSingleAndMultiCellTables(
        string formula,
        string expected
    )
    {
        var value = Eval(ThreeByThreeFixture(), formula);
        var actual = value switch
        {
            ErrorValue error => error.ErrorCode,
            double number => number.ToString(System.Globalization.CultureInfo.InvariantCulture),
            _ => value?.ToString() ?? "null",
        };

        await Assert.That(actual).IsEqualTo(expected);
    }

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
    [Arguments("=INDEX(OFFSET(INDEX(E5:H10,0,1),1,0),1,1)", 6.0)] // inherited 6x1 OFFSET starts at E6
    [Arguments("=ROW(INDEX(E5:H10,0,1))-4", 1.0)] // lifted under a scalar operator
    [Arguments("=SUM(INDEX(E5:H10,0,1):H10)", 180.0)] // ':' range endpoint: bounding box E5:H10
    public async Task ColumnForm_NumericConsumers_MatchTheOracle(string formula, double expected)
    {
        await Assert.That(Num(Eval(IdxFixture(), formula))).IsEqualTo(expected);
    }

    // === Sweep item 37 follow-up, ruling (b): a bare INDEX/OFFSET result under an operator, in an array
    // context, lifts element-wise like the literal range it denotes — the same mechanism item 32 already
    // gave If/Choose (ArrayEvaluation.IsBareReferenceNode's structural arm + the Probe/TryBuildOperand pair
    // reusing WrapScalar), extended to Index and Offset nodes. MySheet has one reading regardless of entry
    // mode (the project's own established convention — see e.g. SUMPRODUCT(-E6:E8) in the divergence
    // probe), which is the oracle's CSE column wherever PLAIN and CSE split.
    [Test]
    [Arguments("=SUMPRODUCT((INDEX(E5:H10,0,1)>6)*1)", 4.0)]
    [Arguments("=SUM((INDEX(E5:H10,0,1)>6)*1)", 4.0)] // oracle PLAIN #VALUE! / CSE 4 — MySheet follows CSE
    [Arguments("=SUMPRODUCT((OFFSET(E5,0,0,6,1)>6)*1)", 4.0)]
    [Arguments("=SUM((OFFSET(E5,0,0,6,1)>6)*1)", 4.0)] // same PLAIN/CSE split, same convention
    public async Task BareIndexOrOffsetUnderAnOperator_LiftsElementwise(
        string formula,
        double expected
    )
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

    // Oracle (Aspose.Cells 26.7.0, Main!AZ5000; PLAIN/CSE): every negative truncated axis is #VALUE! / #VALUE!.
    // The adjacent -0.5 row truncates to zero, so PLAIN/CSE split #VALUE! / 1; bare-cell intersection follows PLAIN.
    [Test]
    [Arguments("=INDEX(A1:C3,-1,0)")]
    [Arguments("=INDEX(A1:C3,-1,1)")]
    [Arguments("=INDEX(A1:C3,0,-1)")]
    [Arguments("=INDEX(A1:C3,1,-1)")]
    [Arguments("=INDEX(A1:A3,-1)")]
    public async Task NegativeIndex_IsValueError(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");

        for (var row = 1; row <= 3; row++)
        for (var column = 1; column <= 3; column++)
        {
            sheet[new CellAddress(column, row).ToId()] = new NumberValue((row - 1) * 3 + column);
        }

        await Assert.That(Eval(workbook, formula)).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task FractionalNegativeIndex_TruncatesToZero()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(4);
        sheet["A3"] = new NumberValue(7);

        await Assert.That(Eval(workbook, "=INDEX(A1:C3,-0.5,1)")).IsEqualTo(ErrorValue.NotValue);
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

    // === Sweep item 37 follow-up, ruling (a): a whole-row/column open base addresses ABSOLUTE grid
    // positions (column A / row 1 origin), not the POPULATED bounding box's own corner
    // (OpenRangeReference.AbsoluteRow/AbsoluteColumn, Index.IndexIntoOpenRange). Was a registered
    // divergence (SUM(INDEX($5:$10,0,5)) #REF! against the oracle's 45) until this ruling closed it.
    [Test]
    public async Task WholeRowOpenBase_AddressesAbsoluteColumns_MatchesTheOracle()
    {
        // Column 5 of the whole row $5:$10 is the ABSOLUTE column E (only E..H happen to be populated),
        // not "the 5th populated column" (there are only 4) — the distinction WholeColumnConsumerTests'
        // Index_WholeColumn_ByAbsolutePosition also pins, over a fixture where the two readings diverge.
        // Expected value changed #REF! -> 45: open-range INDEX now addresses absolute column E.
        await Assert.That(Num(Eval(IdxFixture(), "=SUM(INDEX($5:$10,0,5))"))).IsEqualTo(45.0);
    }

    // === Sweep item 37 follow-up, ruling (b): INDEX under an operator, in an array context, lifts like
    // the literal range it denotes (ArrayEvaluation.WrapScalar, extended to Index/Offset the same way
    // item 32 already extended it to If/Choose). Was a registered divergence (MySheet 1, oracle 3, its
    // denominator comparing a bare INDEX(...) result against "" and 6 scalar-only) until this ruling
    // closed it — this is the divergence probe's row 43, a MYSHEET-CALC-DIVERGENCES.md formula, distinct
    // from the plan's own corpus shape above (CorpusAggregateShape_MatchesTheOracle), which already matched.
    [Test]
    public async Task BareIndexUnderAnOperator_LiftsElementwise_MatchesTheOracle()
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
            // Expected value changed 1 -> 3: the INDEX reference now lifts element-wise under operators.
            .IsEqualTo(3.0);
    }
}
