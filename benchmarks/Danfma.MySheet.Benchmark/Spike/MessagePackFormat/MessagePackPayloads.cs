using Danfma.MySheet.Benchmark.Spike.WholeColumnScale;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.MessagePackFormat;

// The three representative payloads for plans/messagepack-spike.md (size + speed axes). Every formula parses
// into a node the mirror subset covers (values, refs, aggregates, logical, lookup, stat-if, text, scalar
// math, LET) — MirrorConverter throws otherwise, so a payload can never drift outside the measured subset.
public static class MessagePackPayloads
{
    public enum Size
    {
        Small,
        Medium,
        Large,
    }

    public static Workbook Build(Size size) =>
        size switch
        {
            Size.Small => BuildSmall(),
            Size.Medium => BuildMedium(cells: 5_000),
            Size.Large => BuildLarge(dataCells: 100_000),
            _ => throw new ArgumentOutOfRangeException(nameof(size)),
        };

    // (a) Small, fixture-like: ~20 cells mixing values, a cell reference, a binary op and a few formulas
    // spanning categories (SUM aggregate, IF logical, VLOOKUP lookup, ROUND scalar math, CONCAT text, LET).
    private static Workbook BuildSmall()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        for (var r = 1; r <= 5; r++)
        {
            sheet[$"A{r}"] = new Expressions.NumberValue(r * 1.5);
            sheet[$"B{r}"] = new Expressions.StringValue($"label{r}");
        }

        sheet["C1"] = Parse("=A1+A2*2-3^2", sheet);
        sheet["C2"] = Parse("=SUM(A1:A5)", sheet);
        sheet["C3"] = Parse("=AVERAGE(A1:A5)", sheet);
        sheet["C4"] = Parse("=IF(A1>A2, A1, A2)", sheet);
        sheet["C5"] = Parse("=ROUND(A3, 2)", sheet);
        sheet["C6"] = Parse("=CONCAT(B1, B2)", sheet);
        sheet["C7"] = Parse("=VLOOKUP(A1, A1:B5, 2, FALSE)", sheet);
        sheet["C8"] = Parse("=COUNTIF(A1:A5, \">2\")", sheet);
        sheet["C9"] = Parse("=LET(x, A1, x*2)", sheet);
        sheet["C10"] = Parse("=A1:A5", sheet); // a bare range reference cell

        workbook.DefineName("Threshold", Parse("=A1", sheet));

        return workbook;
    }

    // (b) Medium ~5k cells: a data block of numbers + a parallel block of mixed-category formulas, so the
    // payload is a realistic blend of value cells and formula cells (not just one shape).
    private static Workbook BuildMedium(int cells)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Data");

        var dataRows = cells / 2;

        // Value block: A = numbers, B = text.
        for (var r = 1; r <= dataRows; r++)
        {
            sheet[$"A{r}"] = new Expressions.NumberValue(r);
            sheet[$"B{r}"] = new Expressions.StringValue($"k{r}");
        }

        // Formula block: rotate through the categories so every mirror branch is exercised at volume.
        for (var r = 1; r <= dataRows; r++)
        {
            var probe = (r % dataRows) + 1;
            var text = (r % 8) switch
            {
                0 => $"=SUM(A1:A{Cap(r)})",
                1 => $"=AVERAGE(A1:A{Cap(r)})",
                2 => $"=IF(A{r}>A{probe}, A{r}, A{probe})",
                3 => $"=ROUND(A{r}/3, 2)",
                4 => $"=VLOOKUP({probe}, A1:B{dataRows}, 2, FALSE)",
                5 => $"=COUNTIF(A1:A{dataRows}, \">{probe}\")",
                6 => $"=CONCAT(B{r}, B{probe})",
                _ => $"=A{r}*2 + MIN(A1:A{Cap(r)})",
            };

            sheet[$"F{r}"] = Parse(text, sheet);
        }

        return workbook;

        int Cap(int r) => System.Math.Min(r + 10, dataRows);
    }

    // (c) Large ~100k cells: reuse the WholeColumnScale generator (fixed seed) — one big sorted numeric
    // column + companion value/text columns + a formula block of whole-column VLOOKUPs. Every node it emits
    // (NumberValue, StringValue, OpenRangeReference, RangeReference, CellReference, VLookup) is in the subset.
    private static Workbook BuildLarge(int dataCells)
    {
        // dataCells rows × 3 populated columns (A/B/C) ≈ 3·dataCells value cells, plus a formula block.
        var (workbook, _) = WholeColumnScaleData.Build(
            dataCells: dataCells,
            formulaCount: 2_000,
            formula: ScaleFormula.VLookupExact,
            target: ScaleTarget.BigColumn
        );

        return workbook;
    }

    private static Expressions.Expression Parse(string formula, Sheet sheet) =>
        ExpressionParser.Parse(formula, sheet);
}
