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
}
