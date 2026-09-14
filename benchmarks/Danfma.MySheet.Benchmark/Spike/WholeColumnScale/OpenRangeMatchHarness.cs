using System.Diagnostics;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

public static class OpenRangeMatchHarness
{
    private const int DataCells = 500_000;
    private const int FormulaCells = 1_000;

    public static void Run()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        for (var row = 1; row <= DataCells; row++)
        {
            sheet[$"A{row}"] = new NumberValue(row);
        }

        Measure(workbook, sheet, "MATCH", "=MATCH(499999,A:A,0)", "Z");
        Measure(workbook, sheet, "XMATCH", "=XMATCH(499999,A:A)", "Y");
    }

    private static void Measure(
        Workbook workbook,
        Sheet sheet,
        string name,
        string formula,
        string column
    )
    {
        // Warm-up marks the shared open range; the first measured formula builds its snapshot and every
        // remaining distinct formula reuses the same coordinate-aware exact index.
        sheet[$"{column}1"] = ExpressionParser.Parse(formula, sheet);
        _ = workbook.GetCellValue("Sheet1", $"{column}1");

        for (var i = 0; i < FormulaCells; i++)
        {
            sheet[$"{column}{i + 10}"] = ExpressionParser.Parse(formula, sheet);
        }

        var stopwatch = Stopwatch.StartNew();
        for (var i = 0; i < FormulaCells; i++)
        {
            _ = workbook.GetCellValue("Sheet1", $"{column}{i + 10}");
        }
        stopwatch.Stop();

        Console.WriteLine(
            $"{name}: {stopwatch.Elapsed.TotalMilliseconds / FormulaCells:F6} ms/evaluation "
                + $"({DataCells:N0} cells, {FormulaCells:N0} distinct formulas)"
        );
    }
}
