using System.Diagnostics;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

public static class OpenRangeLookupHarness
{
    private const int DataCells = 100_000;
    private const int FormulaCells = 1_000;

    public static void Run()
    {
        MeasureVertical("VLOOKUP exact closed", "Data!A1:B100000", approximate: false);
        MeasureVertical("VLOOKUP approximate closed", "Data!A1:B100000", approximate: true);
        MeasureVertical("VLOOKUP exact open", "Data!A:B", approximate: false);
        MeasureVertical("VLOOKUP approximate open", "Data!A:B", approximate: true);
        MeasureHorizontal("HLOOKUP exact closed", "Data!A1:CVW2", approximate: false);
        MeasureHorizontal("HLOOKUP approximate closed", "Data!A1:CVW2", approximate: true);
        MeasureHorizontal("HLOOKUP exact open", "Data!1:2", approximate: false);
        MeasureHorizontal("HLOOKUP approximate open", "Data!1:2", approximate: true);

        // Retain the existing open-range cache gates beside the table-lookup scenarios.
        var workbook = BuildVerticalWorkbook();
        var sheet = workbook["Formulas"];
        Measure(workbook, sheet, "XLOOKUP", "=XLOOKUP(99999,Data!A:A,Data!B:B)", "B");
        Measure(workbook, sheet, "MATCH", "=MATCH(99999,Data!A:A,0)", "C");
        Measure(workbook, sheet, "XMATCH", "=XMATCH(99999,Data!A:A)", "D");
    }

    private static void MeasureVertical(string name, string table, bool approximate)
    {
        var workbook = BuildVerticalWorkbook();
        Measure(
            workbook,
            workbook["Formulas"],
            name,
            $"=VLOOKUP(99999,{table},2,{(approximate ? "TRUE" : "FALSE")})",
            "A"
        );
    }

    private static Workbook BuildVerticalWorkbook()
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        workbook.Sheets.Add("Formulas");
        for (var row = 1; row <= DataCells; row++)
        {
            data[$"A{row}"] = new NumberValue(row);
            data[$"B{row}"] = new NumberValue(row * 2);
        }

        return workbook;
    }

    private static void MeasureHorizontal(string name, string table, bool approximate)
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add("Data");
        var formulas = workbook.Sheets.Add("Formulas");
        for (var column = 1; column <= DataCells; column++)
        {
            var id = ColumnId(column);
            data[$"{id}1"] = new NumberValue(column);
            data[$"{id}2"] = new NumberValue(column * 2);
        }

        Measure(
            workbook,
            formulas,
            name,
            $"=HLOOKUP(99999,{table},2,{(approximate ? "TRUE" : "FALSE")})",
            "A"
        );
    }

    private static void Measure(
        Workbook workbook,
        Sheet sheet,
        string name,
        string formula,
        string formulaColumn
    )
    {
        for (var index = 1; index <= FormulaCells; index++)
        {
            sheet[$"{formulaColumn}{index}"] = ExpressionParser.Parse(formula, sheet);
        }

        var stopwatch = Stopwatch.StartNew();
        for (var index = 1; index <= FormulaCells; index++)
        {
            _ = workbook.GetCellValue(sheet.Name, $"{formulaColumn}{index}");
        }
        stopwatch.Stop();

        Console.WriteLine(
            $"{name}\t{stopwatch.Elapsed.TotalMilliseconds / FormulaCells:F6}\t"
                + $"{DataCells} cells\t{FormulaCells} distinct formulas"
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
