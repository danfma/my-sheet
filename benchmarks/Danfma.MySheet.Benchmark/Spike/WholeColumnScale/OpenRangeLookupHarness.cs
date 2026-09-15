using System.Diagnostics;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

public static class OpenRangeLookupHarness
{
    private const int DataCells = 500_000;
    private const int FormulaCells = 1_000;

    public static void Run()
    {
        foreach (var hasGaps in new[] { false, true })
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.Add("Sheet1");
            for (var row = 1; row <= DataCells; row++)
            {
                if (hasGaps && row % 7 == 0)
                {
                    continue;
                }

                sheet[$"A{row}"] = new NumberValue(row);
                sheet[$"B{row}"] = new NumberValue(row * 10);
                sheet[$"C{row}"] = new NumberValue(row * 100);
            }

            Console.WriteLine(hasGaps ? "blank every 7th row" : "gap-free column");
            Measure(workbook, sheet, "XLOOKUP(499999,A:A,B:B)", "Z");
            Measure(workbook, sheet, "COUNTIF(XLOOKUP(499999,A:A,B:C),\">0\")", "Y");
            Measure(workbook, sheet, "SUM(XLOOKUP(499999,A:A,B:C))", "X");
            Measure(workbook, sheet, "MATCH(499999,A:A,0)", "W");
            Measure(workbook, sheet, "XMATCH(499999,A:A)", "V");
        }
    }

    private static void Measure(Workbook workbook, Sheet sheet, string expression, string column)
    {
        var formula = $"={expression}";
        sheet[$"{column}1"] = ExpressionParser.Parse(formula, sheet);
        _ = workbook.GetCellValue("Sheet1", $"{column}1");

        for (var index = 0; index < FormulaCells; index++)
        {
            sheet[$"{column}{index + 10}"] = ExpressionParser.Parse(formula, sheet);
        }

        var stopwatch = Stopwatch.StartNew();
        for (var index = 0; index < FormulaCells; index++)
        {
            _ = workbook.GetCellValue("Sheet1", $"{column}{index + 10}");
        }
        stopwatch.Stop();

        Console.WriteLine(
            $"{expression, -44} {stopwatch.Elapsed.TotalMilliseconds / FormulaCells, 12:F6} ms/evaluation"
        );
    }
}
