using System.Globalization;
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// The whole-column function each formula block exercises. Every one is a real Excel function that
// consumes a whole-column reference and, on the current engine, pays an O(N) scan PER formula — so a
// block of F formulas over a column of N cells is O(F*N). See plans/whole-column-performance.md.
public enum ScaleFormula
{
    Match1, // MATCH(x, key:key, 1) — approximate (ascending) match
    Match0, // MATCH(x, key:key, 0) — exact match
    VLookupExact, // VLOOKUP(x, table, 2, FALSE)
    SumIfEqual, // SUMIF(key:key, x)
    CountIfEqual, // COUNTIF(key:key, x)
    Small, // SMALL(key:key, k)
    SumRepeated, // SUM(key:key) — the same aggregate repeated in F cells
}

// Which column the formula block references.
//   BigColumn    → the huge data column (the "consumes the big column" case; Phase 2 territory).
//   NarrowColumn → a tiny column that lives in the SAME big sheet (the "small columns in a big sheet"
//                  case: today every formula still scans ALL keys to find the few in that column, so
//                  it is O(F*N); Phase 1's structural index makes it O(F * narrow) — it collapses).
public enum ScaleTarget
{
    BigColumn,
    NarrowColumn,
}

// Synthetic workbook generator (fixed seed) for the whole-column scale benchmark. One big sorted
// numeric column plus a companion value column (for VLOOKUP), a parallel text column (for the exact
// variant) and a tiny narrow table off to the side. Formula cells reference either the big column or
// the narrow one. Public API only — the benchmark project sees no internals.
internal static class WholeColumnScaleData
{
    // Big data table: keys in column A (row value = row number, so it is sorted ascending), companion
    // values in column B (= row * 10), text keys in column C ("k{row}").
    private const string BigKeyColumn = "A";
    private const string BigTextColumn = "C";
    private const string BigTable = "A:B";

    // Narrow table (few cells) far from A so a full-key scan is forced to walk past the big column to
    // reach it: keys in column K, values in column L, text keys in column M.
    private const string NarrowKeyColumn = "K";
    private const string NarrowTextColumn = "M";
    private const string NarrowTable = "K:L";
    public const int NarrowCount = 16;

    private const string DataSheet = "Data";
    private const string CalcSheet = "Calc";

    public static (Workbook Workbook, string[] FormulaIds) Build(
        int dataCells,
        int formulaCount,
        ScaleFormula formula,
        ScaleTarget target
    )
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add(DataSheet);
        var calc = workbook.Sheets.Add(CalcSheet);

        // Big column: A{r}=r (sorted), B{r}=r*10, C{r}="k{r}".
        for (var r = 1; r <= dataCells; r++)
        {
            data[$"{BigKeyColumn}{r}"] = Number(r);
            data[$"B{r}"] = Number(r * 10);
            data[$"{BigTextColumn}{r}"] = Text($"k{r}");
        }

        // Narrow table: K{r}=r, L{r}=r*10, M{r}="k{r}" for r in 1..NarrowCount.
        for (var r = 1; r <= NarrowCount; r++)
        {
            data[$"{NarrowKeyColumn}{r}"] = Number(r);
            data[$"L{r}"] = Number(r * 10);
            data[$"{NarrowTextColumn}{r}"] = Text($"k{r}");
        }

        var maxKey = target == ScaleTarget.BigColumn ? dataCells : NarrowCount;
        var keyRef = target == ScaleTarget.BigColumn ? $"{BigKeyColumn}:{BigKeyColumn}" : $"{NarrowKeyColumn}:{NarrowKeyColumn}";
        var textRef = target == ScaleTarget.BigColumn ? $"{BigTextColumn}:{BigTextColumn}" : $"{NarrowTextColumn}:{NarrowTextColumn}";
        var tableRef = target == ScaleTarget.BigColumn ? BigTable : NarrowTable;

        var formulaIds = new string[formulaCount];

        for (var i = 0; i < formulaCount; i++)
        {
            // A varying but in-range probe so the lookups actually hit (and vary cell to cell).
            var probe = (i % maxKey) + 1;
            var probeText = $"\"k{probe}\"";
            var probeNumber = probe.ToString(CultureInfo.InvariantCulture);

            var text = formula switch
            {
                ScaleFormula.Match1 => $"=MATCH({probeNumber},Data!{keyRef},1)",
                ScaleFormula.Match0 => $"=MATCH({probeText},Data!{textRef},0)",
                ScaleFormula.VLookupExact => $"=VLOOKUP({probeNumber},Data!{tableRef},2,FALSE)",
                ScaleFormula.SumIfEqual => $"=SUMIF(Data!{keyRef},{probeNumber})",
                ScaleFormula.CountIfEqual => $"=COUNTIF(Data!{keyRef},{probeNumber})",
                ScaleFormula.Small => $"=SMALL(Data!{keyRef},{probeNumber})",
                ScaleFormula.SumRepeated => $"=SUM(Data!{keyRef})",
                _ => throw new ArgumentOutOfRangeException(nameof(formula)),
            };

            var id = $"A{i + 1}";
            calc[id] = ExpressionParser.Parse(text, calc);
            formulaIds[i] = id;
        }

        return (workbook, formulaIds);
    }

    // Evaluates every formula cell once (values are memoized, so this measures one full O(F*N) pass),
    // returning a checksum so the JIT cannot elide the work.
    public static double EvaluateAll(Workbook workbook, string[] formulaIds)
    {
        var checksum = 0.0;

        foreach (var id in formulaIds)
        {
            var value = workbook.GetCellValue(CalcSheet, id);
            if (value.AsObject() is double number)
            {
                checksum += number;
            }
        }

        return checksum;
    }

    private static Expressions.Expression Number(double value) => new Expressions.NumberValue(value);

    private static Expressions.Expression Text(string value) => new Expressions.StringValue(value);
}
