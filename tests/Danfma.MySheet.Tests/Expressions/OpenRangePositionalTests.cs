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
/// result follow automatically once <c>INDEX</c> is correct. The performance model is unchanged: every
/// path here still visits only POPULATED cells through the structural index
/// (<see cref="OpenRangeReference.PopulatedCells"/>), never the whole grid.
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

    // === Absolute coordinates over an open base ============================================================

    [Test]
    public async Task Match_OverAWholeRow_ReturnsTheAbsoluteColumnPosition()
    {
        // "x" sits at C4 — absolute column C, the 3rd column — not position 1 (the only POPULATED cell
        // MATCH would otherwise count as the first).
        await Assert.That(Num(Eval(CorpusFixture(), "=MATCH(\"x\",$4:$4,0)"))).IsEqualTo(3.0);
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
