using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ClosedXML.Excel;
using MySheet.Expressions;
using MySheet.Parsing;
using static MySheet.Expressions.Expression;

namespace MySheet.Benchmark;

[MemoryDiagnoser, GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class SheetBenchmarks
{
    private const string Formula = "=SUM(A1:A3) + AVERAGE(B1,B2)*2 - 3^2";

    private Workbook _mySheetWorkbook = null!;
    private XLWorkbook _closedXmlWorkbook = null!;
    private Sheet _parseSheet = null!;

    [GlobalSetup]
    public void Setup()
    {
        _mySheetWorkbook = CreateMySheetWorkbook();
        _closedXmlWorkbook = CreateClosedXmlWorkbook();
        _parseSheet = _mySheetWorkbook["Sheet1"];

        // Comparison/conditional cells live here (not in CreateMySheetWorkbook) so the InMemory
        // benchmark keeps measuring a pure build, free of parsing.
        _parseSheet["B1"] = ExpressionParser.Parse("=A1>A2", _parseSheet);
        _parseSheet["B2"] = ExpressionParser.Parse("=\"hello\"=\"HELLO\"", _parseSheet);
        _parseSheet["B3"] = ExpressionParser.Parse("=IF(A1>A2, A1, A2)", _parseSheet);
    }

    [Benchmark, BenchmarkCategory("Parse")]
    public Expression MySheetParse()
    {
        return ExpressionParser.Parse(Formula, _parseSheet);
    }

    [Benchmark, BenchmarkCategory("SheetInMemory")]
    public MySheet.Workbook MySheetInMemory()
    {
        return CreateMySheetWorkbook();
    }

    [Benchmark, BenchmarkCategory("SheetInMemory")]
    public XLWorkbook ClosedXmlInMemory()
    {
        return CreateClosedXmlWorkbook();
    }

    [Benchmark, BenchmarkCategory("Compute")]
    public double? MySheetSum()
    {
        return _mySheetWorkbook["Sheet1"]["A3"].Compute(_mySheetWorkbook) as double?;
    }

    [Benchmark, BenchmarkCategory("Compute")]
    public double? ClosedXmlSum()
    {
        return _closedXmlWorkbook.Worksheet("Sheet1").Cell("A3").GetDouble();
    }

    [Benchmark, BenchmarkCategory("Comparison")]
    public object? MySheetComparison()
    {
        return _mySheetWorkbook["Sheet1"]["B1"].Compute(_mySheetWorkbook);
    }

    [Benchmark, BenchmarkCategory("Comparison")]
    public bool ClosedXmlComparison()
    {
        return _closedXmlWorkbook.Worksheet("Sheet1").Cell("B1").GetBoolean();
    }

    [Benchmark, BenchmarkCategory("TextEquality")]
    public object? MySheetTextEquality()
    {
        return _mySheetWorkbook["Sheet1"]["B2"].Compute(_mySheetWorkbook);
    }

    [Benchmark, BenchmarkCategory("TextEquality")]
    public bool ClosedXmlTextEquality()
    {
        return _closedXmlWorkbook.Worksheet("Sheet1").Cell("B2").GetBoolean();
    }

    [Benchmark, BenchmarkCategory("Conditional")]
    public double? MySheetConditional()
    {
        return _mySheetWorkbook["Sheet1"]["B3"].Compute(_mySheetWorkbook) as double?;
    }

    [Benchmark, BenchmarkCategory("Conditional")]
    public double? ClosedXmlConditional()
    {
        return _closedXmlWorkbook.Worksheet("Sheet1").Cell("B3").GetDouble();
    }

    private static Workbook CreateMySheetWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = Number(1);
        sheet["A2"] = Number(2);
        sheet["A3"] = Sum(Cell("A1", sheet), Cell("A2", sheet));

        return workbook;
    }

    private static XLWorkbook CreateClosedXmlWorkbook()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");

        sheet.Cell("A1").Value = 1;
        sheet.Cell("A2").Value = 2;
        sheet.Cell("A3").FormulaA1 = "=SUM(A1,A2)";
        sheet.Cell("B1").FormulaA1 = "=A1>A2";
        sheet.Cell("B2").FormulaA1 = "=\"hello\"=\"HELLO\"";
        sheet.Cell("B3").FormulaA1 = "=IF(A1>A2,A1,A2)";

        return workbook;
    }
}
