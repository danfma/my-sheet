using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// Sweep item 37 follow-up, controller ruling (a): a positional function's index or position over an OPEN
/// range (<c>$4:$4</c>, <c>$5:$1000</c>, <c>A:A</c>, …) counts from the range's own declared origin —
/// column A / row 1 when that side is open — never the POPULATED bounding box's own corner
/// (<c>OpenRangeReference.ToBoundedRange</c>). <c>MATCH</c>/<c>XMATCH</c> return the absolute position;
/// <c>INDEX(open,r,c)</c> addresses absolute rows/columns; <c>ROW</c>/<c>COLUMN</c>/<c>COUNTA</c> of the
/// result follow automatically once <c>INDEX</c> is correct. <c>INDEX</c>/<c>OFFSET</c> translate coordinates
/// arithmetically; <c>MATCH</c>/<c>XMATCH</c> visit only POPULATED cells through the structural index and
/// retain their source positions. No path materializes the whole grid.
///
/// <para>Oracle: Aspose.Cells 26.7.0, one formula per workbook, 2026-09-14, PLAIN and CSE agreeing on
/// every row measured. The "corpus" fixture mirrors the downstream consumer's own idiom:
/// <c>Main!C4</c> = "x" (a header cell in a whole-ROW range), <c>C5:C8</c> = "a" / "skip" / "b" / "c"
/// (data below it, in a whole-ROW-bounded / whole-COLUMN-open table), <c>D5</c> = 1.</para>
/// </summary>
public class OpenRangePositionalTests
{
    private static Workbook CorpusFixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C4"] = new StringValue("x");
        sheet["C5"] = new StringValue("a");
        sheet["C6"] = new StringValue("skip");
        sheet["C7"] = new StringValue("b");
        sheet["C8"] = new StringValue("c");
        sheet["D5"] = new NumberValue(1);

        return workbook;
    }

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

    private static object? Eval(Workbook workbook, string formula, string cell = "Z100")
    {
        var sheet = workbook["Main"];
        sheet[cell] = ExpressionParser.Parse(formula, sheet);

        return workbook.GetCellValue("Main", cell).AsObject();
    }

    private static double? Num(object? value) => value as double?;

    private static object? EvalAdmitted(Workbook workbook, string formula)
    {
        // The first distinct formula cell records the cache-admission marker; the second builds and reads the
        // snapshot. Keeping the formula cells off the lookup range avoids including either formula in it.
        _ = Eval(workbook, formula, "Z100");
        return Eval(workbook, formula, "AA100");
    }

    private static void FillRowForCacheAdmission(Sheet sheet)
    {
        for (var column = 4; column <= 300; column++)
        {
            sheet[$"{ColumnAddress(column)}4"] = new NumberValue(0);
        }
    }

    private static void FillColumnForCacheAdmission(Sheet sheet, int startRow, double value)
    {
        for (var row = startRow; row < startRow + 257; row++)
        {
            sheet[$"C{row}"] = new NumberValue(value);
        }
    }

    private static string ColumnAddress(int column)
    {
        var address = string.Empty;
        while (column > 0)
        {
            column--;
            address = (char)('A' + column % 26) + address;
            column /= 26;
        }

        return address;
    }

    private static RangeSnapshot? SnapshotForOpenRange(Workbook workbook)
    {
        var cache =
            typeof(Workbook)
                .GetField(
                    "_rangeCache",
                    System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.NonPublic
                )!
                .GetValue(workbook) as System.Collections.IEnumerable;

        if (cache is null)
        {
            return null;
        }

        foreach (var item in cache)
        {
            var itemType = item!.GetType();
            if (itemType.GetProperty("Key")!.GetValue(item) is not OpenRangeReference)
            {
                continue;
            }

            var entry = itemType.GetProperty("Value")!.GetValue(item)!;
            return (RangeSnapshot?)entry.GetType().GetProperty("Snapshot")!.GetValue(entry);
        }

        return null;
    }

    private static bool ExactIndexWasBuilt(RangeSnapshot snapshot)
    {
        var field = snapshot
            .GetType()
            .GetField(
                "_exact",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic
            )!;
        var lazy = field.GetValue(snapshot)!;
        return (bool)field.FieldType.GetProperty("IsValueCreated")!.GetValue(lazy)!;
    }

    // === Absolute coordinates over an open base ============================================================

    [Test]
    [Arguments("=XLOOKUP(2,A:A,B:B)", 30.0, 20.0)]
    [Arguments("=XLOOKUP(3,A:A,B:B)", 50.0, 30.0)]
    [Arguments("=SUM(XLOOKUP(3,A:A,B:C))", 50.0, 330.0)]
    [Arguments("=XLOOKUP(2,1:1,2:2)", 9.0, 7.0)]
    [Arguments("=MATCH(3,A:A,0)", 5.0, 5.0)]
    [Arguments("=XMATCH(3,A:A)", 5.0, 5.0)]
    public async Task OpenLookupFunctions_UsePopulatedCellCoordinates(
        string formula,
        double expected,
        double valueBeforeFix
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["A2"] = new NumberValue(1);
        sheet["A3"] = new NumberValue(2);
        sheet["A5"] = new NumberValue(3);
        for (var row = 1; row <= 5; row++)
        {
            sheet[$"B{row}"] = new NumberValue(row * 10);
        }

        sheet["C1"] = new NumberValue(1);
        sheet["E1"] = new NumberValue(2);
        sheet["C2"] = new NumberValue(7);
        sheet["E2"] = new NumberValue(9);

        // Aspose.Cells 26.7.0 PLAIN/CSE agree on expected. valueBeforeFix records the old MySheet value.
        _ = valueBeforeFix;
        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
    }

    // === Cached coordinates over an open base ===============================================================

    [Test]
    public async Task Match_OverAnAdmittedWholeRow_ReturnsTheSourceColumnPosition()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C4"] = new NumberValue(30);
        FillRowForCacheAdmission(sheet);

        // Aspose.Cells 26.7.0 PLAIN/CSE: 3 / 3. The second distinct formula cell uses the admitted snapshot.
        await Assert.That(Num(EvalAdmitted(workbook, "=MATCH(30,$4:$4,0)"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task XMatch_OverAnAdmittedWholeRow_ReturnsTheSourceColumnPosition()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C4"] = new NumberValue(30);
        FillRowForCacheAdmission(sheet);

        // Aspose.Cells 26.7.0 PLAIN/CSE: 3 / 3. The second distinct formula cell uses the admitted snapshot.
        await Assert.That(Num(EvalAdmitted(workbook, "=XMATCH(30,$4:$4)"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task Match_AscendingOverAnAdmittedWholeColumn_ReturnsTheSourceRowPosition()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C3"] = new NumberValue(20);
        sheet["C500"] = new NumberValue(30);
        sheet["C900"] = new NumberValue(40);
        FillColumnForCacheAdmission(sheet, 901, 40);

        // Aspose.Cells 26.7.0 PLAIN/CSE: 500 / 500. The second distinct formula cell uses the admitted snapshot.
        await Assert.That(Num(EvalAdmitted(workbook, "=MATCH(35,C:C,1)"))).IsEqualTo(500.0);
    }

    [Test]
    public async Task Match_DescendingOverAnAdmittedWholeColumn_ReturnsTheSourceRowPosition()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C3"] = new NumberValue(40);
        sheet["C500"] = new NumberValue(30);
        sheet["C900"] = new NumberValue(20);
        FillColumnForCacheAdmission(sheet, 901, 20);

        // Aspose.Cells 26.7.0 PLAIN/CSE: 3 / 3. The second distinct formula cell uses the admitted snapshot.
        await Assert.That(Num(EvalAdmitted(workbook, "=MATCH(35,C:C,-1)"))).IsEqualTo(3.0);
    }

    [Test]
    public async Task XMatch_OverAnAdmittedWholeColumn_ReturnsTheSparseSourceRowPosition()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C7"] = new NumberValue(30);
        FillColumnForCacheAdmission(sheet, 8, 0);

        // Aspose.Cells 26.7.0 PLAIN/CSE: 7 / 7. The second distinct formula cell uses the admitted snapshot.
        await Assert.That(Num(EvalAdmitted(workbook, "=XMATCH(30,C:C)"))).IsEqualTo(7.0);
    }

    [Test]
    [Arguments("=XMATCH(30,C:C)", 7.0)]
    [Arguments("=XLOOKUP(30,C:C,D:D)", 70.0)]
    public async Task ExactOpenLookup_OverAnAdmittedSparseColumn_UsesTheSnapshotIndex(
        string formula,
        double expected
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        sheet["C7"] = new NumberValue(30);
        sheet["D7"] = new NumberValue(70);
        FillColumnForCacheAdmission(sheet, 8, 0);

        await Assert.That(Num(EvalAdmitted(workbook, formula))).IsEqualTo(expected);
        var snapshot = SnapshotForOpenRange(workbook);
        await Assert.That(snapshot is not null).IsTrue();
        await Assert.That(ExactIndexWasBuilt(snapshot!)).IsTrue();
    }

    [Test]
    public async Task Match_OverAWholeRow_ReturnsTheAbsoluteColumnPosition()
    {
        // "x" sits at C4 — absolute column C, the 3rd column — not position 1 (the only POPULATED cell
        // MATCH would otherwise count as the first).
        await Assert.That(Num(Eval(CorpusFixture(), "=MATCH(\"x\",$4:$4,0)"))).IsEqualTo(3.0);
    }

    [Test]
    [Arguments("=XMATCH(30,$4:$4,0,1)", 3.0)]
    [Arguments("=XMATCH(30,$4:$4,0,-1)", 7.0)]
    [Arguments("=XMATCH(30,$4:$4,0,2)", 5.0)]
    [Arguments("=XMATCH(30,$4:$4,0,-2)", 5.0)]
    [Arguments("=XMATCH(35,$4:$4,-1,1)", 5.0)]
    [Arguments("=XMATCH(35,$4:$4,1,1)", 7.0)]
    [Arguments("=XMATCH(\"a*\",$4:$4,2,1)", 3.0)]
    [Arguments("=XMATCH(\"a*\",$4:$4,2,-1)", 7.0)]
    public async Task XMatch_OverAWholeRow_ReturnsTheAbsoluteColumnPosition(
        string formula,
        double expected
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        if (formula.Contains("a*"))
        {
            sheet["C4"] = new StringValue("alpha");
            sheet["E4"] = new StringValue("beta");
            sheet["G4"] = new StringValue("atom");
        }
        else if (
            formula.Contains("XMATCH(35")
            || formula.EndsWith(",2)")
            || formula.EndsWith(",-2)")
        )
        {
            sheet["C4"] = new NumberValue(20);
            sheet["E4"] = new NumberValue(30);
            sheet["G4"] = new NumberValue(40);
        }
        else
        {
            sheet["C4"] = new NumberValue(30);
            sheet["E4"] = new NumberValue(25);
            sheet["G4"] = new NumberValue(30);
        }

        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=XMATCH(30,C:C,0,1)", 3.0)]
    [Arguments("=XMATCH(30,C:C,0,-1)", 7.0)]
    [Arguments("=XMATCH(35,C:C,-1,1)", 5.0)]
    [Arguments("=XMATCH(35,C:C,1,1)", 7.0)]
    [Arguments("=XMATCH(\"a*\",C:C,2,1)", 3.0)]
    [Arguments("=XMATCH(\"a*\",C:C,2,-1)", 7.0)]
    public async Task XMatch_OverAWholeColumn_ReturnsTheAbsoluteRowPosition(
        string formula,
        double expected
    )
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        if (formula.Contains("a*"))
        {
            sheet["C3"] = new StringValue("alpha");
            sheet["C5"] = new StringValue("beta");
            sheet["C7"] = new StringValue("atom");
        }
        else if (formula.Contains("XMATCH(35"))
        {
            sheet["C3"] = new NumberValue(20);
            sheet["C5"] = new NumberValue(30);
            sheet["C7"] = new NumberValue(40);
        }
        else
        {
            sheet["C3"] = new NumberValue(30);
            sheet["C5"] = new NumberValue(25);
            sheet["C7"] = new NumberValue(30);
        }

        await Assert.That(Num(Eval(workbook, formula))).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=COLUMN(INDEX($5:$1000,0,3))", 3.0)]
    [Arguments("=COUNTA(INDEX($5:$1000,0,3))", 4.0)]
    public async Task Index_OverAWholeRowBase_AddressesTheAbsoluteColumn(
        string formula,
        double expected
    )
    {
        await Assert.That(Num(Eval(CorpusFixture(), formula))).IsEqualTo(expected);
    }

    [Test]
    public async Task Index_OverAWholeRowBase_NonZero_ReadsTheAbsoluteCell()
    {
        // row 1 of $5:$1000 is absolute row 5 (the range's own declared top); column 3 is absolute column
        // C — C5 = "a" — not "the 3rd populated column" (there is only one: C).
        await Assert.That(Eval(CorpusFixture(), "=INDEX($5:$1000,1,3)") as string).IsEqualTo("a");
    }

    [Test]
    [Arguments("=COUNTA(INDEX($5:$10,0,5))", 6.0)]
    [Arguments("=COLUMN(INDEX($5:$10,0,5))", 5.0)]
    public async Task Index_OverAWholeRowBase_IdxGrid_AddressesTheAbsoluteColumn(
        string formula,
        double expected
    )
    {
        await Assert.That(Num(Eval(IdxFixture(), formula))).IsEqualTo(expected);
    }

    // === The corpus's own AGGREGATE idiom: MATCH locates the absolute header column, INDEX(0-row) reads
    // that whole column as a reference, and the fraction's denominator (ruling (b): a bare INDEX(...)
    // result lifts under a comparison) excludes the "skip" sentinel — 3 valid data points {1,3,4}, once
    // for each row/AGGREGATE-k combination the corpus formula is built with.

    [Test]
    public async Task CorpusAggregate_KEqualsOne_ViaNestedIndexRowSubtraction()
    {
        var workbook = CorpusFixture();

        await Assert
            .That(
                Num(
                    Eval(
                        workbook,
                        "=IFERROR(AGGREGATE(15,6,(ROW(INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0)))-ROW(INDEX(INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0)),1,1))+1)/((INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0))<>\"\")*(INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0))<>\"skip\")),ROWS($B$2:B2)),\"\")"
                    )
                )
            )
            .IsEqualTo(1.0);
    }

    [Test]
    [Arguments(2, 3.0)]
    [Arguments(3, 4.0)]
    public async Task CorpusAggregate_KEqualsTwoOrThree_ViaLiteralRowOffset(int k, double expected)
    {
        var workbook = CorpusFixture();

        await Assert
            .That(
                Num(
                    Eval(
                        workbook,
                        $"=IFERROR(AGGREGATE(15,6,(ROW(INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0)))-4)/((INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0))<>\"\")*(INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0))<>\"skip\")),{k}),\"\")"
                    )
                )
            )
            .IsEqualTo(expected);
    }

    [Test]
    public async Task CorpusAggregate_Sumproduct_CountsTheNonBlankColumn()
    {
        // Oracle 4 (a, skip, b, c are all non-blank) — the divergence probe's stale "doc" column claims 3
        // (from ~/MYSHEET-CALC-DIVERGENCES.md, not re-verified against 26.7.0); re-measured directly
        // against Aspose.Cells 26.7.0 here rather than trusted.
        await Assert
            .That(
                Num(
                    Eval(
                        CorpusFixture(),
                        "=SUMPRODUCT((INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0))<>\"\")*1)"
                    )
                )
            )
            .IsEqualTo(4.0);
    }

    [Test]
    public async Task CorpusAggregate_Countifs_CountsTheNonBlankColumn()
    {
        // Same re-measurement as the SUMPRODUCT row above: oracle 4, not the stale doc's 3.
        await Assert
            .That(
                Num(
                    Eval(
                        CorpusFixture(),
                        "=COUNTIFS(INDEX($5:$1000,0,MATCH(\"x\",$4:$4,0)),\"<>\")"
                    )
                )
            )
            .IsEqualTo(4.0);
    }
}
