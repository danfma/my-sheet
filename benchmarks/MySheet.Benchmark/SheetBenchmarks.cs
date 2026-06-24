using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ClosedXML.Excel;
using static MySheet.Expressions.Expression;

namespace MySheet.Benchmark;

[MemoryDiagnoser, GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class SheetBenchmarks
{
    private Workbook _mySheetWorkbook = null!;
    private XLWorkbook _closedXmlWorkbook = null!;

    [GlobalSetup]
    public void Setup()
    {
        _mySheetWorkbook = CreateMySheetWorkbook();
        _closedXmlWorkbook = CreateClosedXmlWorkbook();
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

        return workbook;
    }
}
