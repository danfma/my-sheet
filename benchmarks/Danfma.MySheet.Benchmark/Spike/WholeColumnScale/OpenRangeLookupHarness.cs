using System.Diagnostics;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

public static class OpenRangeLookupHarness
{
    private const int DataCells = 500_000;
    private const int TableCells = 100_000;
    private const int FormulaCells = 1_000;

    public static void Run(bool steady = false)
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
            Measure(workbook, sheet, "XLOOKUP(499999,A:A,B:B)", "Z", steady);
            Measure(workbook, sheet, "COUNTIF(XLOOKUP(499999,A:A,B:C),\">0\")", "Y", steady);
            Measure(workbook, sheet, "SUM(XLOOKUP(499999,A:A,B:C))", "X", steady);
            Measure(workbook, sheet, "MATCH(499999,A:A,0)", "W", steady);
            Measure(workbook, sheet, "XMATCH(499999,A:A)", "V", steady);
        }

        Console.WriteLine("table lookups (100k)");
        var vertical = new Workbook();
        var formulas = vertical.Sheets.Add("Sheet1");
        var data = vertical.Sheets.Add("Data");
        for (var row = 1; row <= TableCells; row++)
        {
            data[$"A{row}"] = new NumberValue(row);
            data[$"B{row}"] = new NumberValue(row * 2);
            data[$"C{row}"] = new StringValue($"k{row:D6}");
            data[$"D{row}"] = new NumberValue(row * 3);
        }

        Measure(vertical, formulas, "VLOOKUP(99999,Data!A1:B100000,2,FALSE)", "Z", steady);
        Measure(vertical, formulas, "VLOOKUP(99999,Data!A1:B100000,2,TRUE)", "Y", steady);
        Measure(vertical, formulas, "VLOOKUP(99999,Data!A:B,2,FALSE)", "X", steady);
        Measure(vertical, formulas, "VLOOKUP(99999,Data!A:B,2,TRUE)", "W", steady);
        Measure(vertical, formulas, "VLOOKUP(\"k099999\",Data!C:D,2,FALSE)", "V", steady);
        Measure(vertical, formulas, "VLOOKUP(\"k09999*\",Data!C:D,2,FALSE)", "U", steady);

        var horizontal = new Workbook();
        var hFormulas = horizontal.Sheets.Add("Sheet1");
        var hData = horizontal.Sheets.Add("Data");
        for (var column = 1; column <= TableCells; column++)
        {
            var id = ColumnId(column);
            hData[$"{id}1"] = new NumberValue(column);
            hData[$"{id}2"] = new NumberValue(column * 2);
            hData[$"{id}3"] = new StringValue($"k{column:D6}");
            hData[$"{id}4"] = new NumberValue(column * 3);
        }

        Measure(horizontal, hFormulas, "HLOOKUP(99999,Data!A1:EQXD2,2,FALSE)", "Z", steady);
        Measure(horizontal, hFormulas, "HLOOKUP(99999,Data!A1:EQXD2,2,TRUE)", "Y", steady);
        Measure(horizontal, hFormulas, "HLOOKUP(99999,Data!1:2,2,FALSE)", "X", steady);
        Measure(horizontal, hFormulas, "HLOOKUP(99999,Data!1:2,2,TRUE)", "W", steady);
        Measure(horizontal, hFormulas, "HLOOKUP(\"k099999\",Data!3:4,2,FALSE)", "V", steady);
        Measure(horizontal, hFormulas, "HLOOKUP(\"k09999*\",Data!3:4,2,FALSE)", "U", steady);
    }

    private static void Measure(
        Workbook workbook,
        Sheet sheet,
        string expression,
        string column,
        bool steady
    )
    {
        var formula = $"={expression}";
        sheet[$"{column}1"] = ExpressionParser.Parse(formula, sheet);
        _ = workbook.GetCellValue("Sheet1", $"{column}1");
        if (steady)
        {
            sheet[$"{column}2"] = ExpressionParser.Parse(formula, sheet);
            _ = workbook.GetCellValue("Sheet1", $"{column}2");
        }

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

    private static string ColumnId(int column)
    {
        Span<char> buffer = stackalloc char[8];
        var index = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--index] = (char)('A' + (column % 26));
            column /= 26;
        }

        return new string(buffer[index..]);
    }
}
