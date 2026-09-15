using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

public class StructuredReferenceColumnSpanTests
{
    // Aspose.Cells 26.7.0, one formula per workbook at Main!AZ5000, PLAIN/CSE identical.
    // Fixture: Data!Tabela1=A1:D4, headers Item/Valor/Qtd/Sales Amount, numeric rows
    // (1,10,100,1000), (2,20,200,2000), (3,30,300,3000); Main!A1:A3=1,2,3.
    [Test]
    [Arguments("=SUM(Tabela1[[Valor]:[Qtd]])", "660")]
    [Arguments("=COUNTIF(Tabela1[[Valor]:[Qtd]],\">0\")", "6")]
    [Arguments("=ROWS(Tabela1[[Valor]:[Qtd]])", "3")]
    [Arguments("=COLUMNS(Tabela1[[Valor]:[Qtd]])", "2")]
    [Arguments("=INDEX(Tabela1[[Valor]:[Qtd]],2,2)", "200")]
    [Arguments("=COUNTIF(XLOOKUP(2,A1:A3,Tabela1[[Valor]:[Qtd]]),\">0\")", "2")]
    [Arguments("=SUM(XLOOKUP(2,A1:A3,Tabela1[[Valor]:[Qtd]]))", "220")]
    [Arguments("=SUM(Tabela1[[Item]:[Sales Amount]])", "6666")]
    [Arguments("=SUM(Tabela1[[Valor]:[Valor]])", "60")]
    [Arguments("=SUM(Tabela1[[Qtd]:[Valor]])", "660")]
    [Arguments("=SUM(Tabela1[[Valor]:[Nope]])", "#REF!")]
    [Arguments("=SUM(Tabela1[[Qtd]:[Sales Amount]])", "6600")]
    [Arguments("=SUM(Tabela1[[O''Brien]:[Qtd]])", "660")]
    [Arguments("=SUM(Tabela1[[#Data],[Valor]:[Qtd]])", "660")]
    [Arguments("=SUM(Tabela1[[#Headers],[Valor]:[Qtd]])", "0")]
    [Arguments("=SUM(Tabela1[[#Totals],[Valor]:[Qtd]])", "#REF!")]
    [Arguments("=SUM(Tabela1[[#All],[Valor]:[Qtd]])", "660")]
    [Arguments("=SUM(Tabela1[[#Headers],[#Data],[Valor]:[Qtd]])", "660")]
    [Arguments("=SUM(INDIRECT(\"Tabela1[[Valor]:[Qtd]]\"))", "660")]
    [Arguments("=AREAS((Tabela1[[Valor]:[Qtd]],Main!A1:A3))", "2")]
    public async Task DataTable_MatchesTheMeasuredOracle(string formula, string expected) =>
        await Assert.That(Evaluate(formula, TableShape.Data)).IsEqualTo(expected);

    [Test]
    [Arguments("=SUM(Tabela1[[Valor]:[Qtd]])", "660")]
    [Arguments("=COUNTIF(Tabela1[[Valor]:[Qtd]],\">0\")", "6")]
    [Arguments("=ROWS(Tabela1[[Valor]:[Qtd]])", "3")]
    [Arguments("=COLUMNS(Tabela1[[Valor]:[Qtd]])", "2")]
    [Arguments("=INDEX(Tabela1[[Valor]:[Qtd]],2,2)", "200")]
    [Arguments("=SUM(Tabela1[[#Totals],[Valor]:[Qtd]])", "0")]
    [Arguments("=SUM(Tabela1[[#All],[Valor]:[Qtd]])", "660")]
    public async Task TotalsTable_MatchesTheMeasuredOracle(string formula, string expected) =>
        await Assert.That(Evaluate(formula, TableShape.Totals)).IsEqualTo(expected);

    [Test]
    [Arguments("=SUM(Tabela1[[Valor]:[Qtd]])", "0")]
    [Arguments("=COUNTIF(Tabela1[[Valor]:[Qtd]],\">0\")", "0")]
    [Arguments("=ROWS(Tabela1[[Valor]:[Qtd]])", "0")]
    [Arguments("=COLUMNS(Tabela1[[Valor]:[Qtd]])", "2")]
    [Arguments("=INDEX(Tabela1[[Valor]:[Qtd]],2,2)", "#REF!")]
    [Arguments("=SUM(Tabela1[[#Totals],[Valor]:[Qtd]])", "#REF!")]
    [Arguments("=SUM(Tabela1[[#All],[Valor]:[Qtd]])", "0")]
    public async Task HeaderOnlyTable_MatchesTheMeasuredOracle(string formula, string expected) =>
        await Assert.That(Evaluate(formula, TableShape.HeaderOnly)).IsEqualTo(expected);

    [Test]
    [Arguments("=SUM(Tabela1[[Valor]:[Qtd]])", "SUM(Tabela1[[Valor]:[Qtd]])")]
    [Arguments("=SUM(Tabela1[[Qtd]:[Valor]])", "SUM(Tabela1[[Qtd]:[Valor]])")]
    [Arguments("=SUM(Tabela1[[Valor] : [Qtd]])", "SUM(Tabela1[[Valor]:[Qtd]])")]
    [Arguments(
        "=SUM(Tabela1[[#Headers],[#Data],[Valor]:[Qtd]])",
        "SUM(Tabela1[[#Headers],[#Data],[Valor]:[Qtd]])"
    )]
    [Arguments("=SUM(Tabela1[[O''Brien]:[Qtd]])", "SUM(Tabela1[[O''Brien]:[Qtd]])")]
    public async Task FormulaText_RoundTripsLikeTheOracle(string formula, string expected)
    {
        var workbook = Fixture(TableShape.Data, escaped: formula.Contains("O''Brien"));
        var main = workbook.Sheets["Main"];
        var expression = ExpressionParser.Parse(formula, main);

        await Assert.That(expression.ToFormula(main.Name)).IsEqualTo(expected);
    }

    [Test]
    public async Task CurrentRowSpan_RemainsOutOfScope()
    {
        var workbook = Fixture(TableShape.Data);
        var data = workbook.Sheets["Data"];
        var error = Assert.Throws<ParseException>(() =>
            ExpressionParser.Parse("=SUM(Tabela1[@[Valor]:[Qtd]])", data)
        );

        await Assert.That(error!.Kind).IsEqualTo(ParseErrorKind.UnsupportedStructuredReference);
    }

    [Test]
    public async Task ColumnSpan_SurvivesMemoryPack()
    {
        Expression reference = new TableReference(
            "Tabela1",
            "Valor",
            TableArea.HeadersAndData,
            "Qtd"
        );

        var clone = MemoryPackSerializer.Deserialize<Expression>(
            MemoryPackSerializer.Serialize(reference)
        );

        await Assert.That(clone).IsEqualTo(reference);
        await Assert
            .That(clone!.ToFormula("Main"))
            .IsEqualTo("Tabela1[[#Headers],[#Data],[Valor]:[Qtd]]");
    }

    [Test]
    public async Task ReversedSpan_NormalizesItsResolvedEndpoints()
    {
        var reference = new TableReference("Tabela1", "Qtd", TableArea.Data, "Valor");

        var ok = reference.TryResolve(Fixture(TableShape.Data), out var range, out _);

        await Assert.That(ok).IsTrue();
        await Assert.That(range).IsEqualTo(new RangeReference("B2", "C4", "Data"));
    }

    private static string Evaluate(string formula, TableShape shape)
    {
        var workbook = Fixture(shape, escaped: formula.Contains("O''Brien"));
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        var value = workbook.GetCellValue("Main", "AZ5000");

        return value.TryGetError(out var error)
            ? error.ToString()
            : value.ToDouble().ToString(CultureInfo.InvariantCulture);
    }

    private static Workbook Fixture(TableShape shape, bool escaped = false)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        var data = workbook.Sheets.Add("Data");
        string[] headers = ["Item", escaped ? "O'Brien" : "Valor", "Qtd", "Sales Amount"];

        for (var row = 1; row <= 3; row++)
        {
            main[$"A{row}"] = new NumberValue(row);
        }

        for (var column = 0; column < headers.Length; column++)
        {
            data[$"{(char)('A' + column)}1"] = new StringValue(headers[column]);
        }

        if (shape is not TableShape.HeaderOnly)
        {
            for (var row = 1; row <= 3; row++)
            {
                data[$"A{row + 1}"] = new NumberValue(row);
                data[$"B{row + 1}"] = new NumberValue(row * 10);
                data[$"C{row + 1}"] = new NumberValue(row * 100);
                data[$"D{row + 1}"] = new NumberValue(row * 1000);
            }
        }

        workbook.DefineTable(
            "Tabela1",
            "Data",
            shape is TableShape.Totals ? "A1:D5"
                : shape is TableShape.HeaderOnly ? "A1:D1"
                : "A1:D4",
            headers,
            hasTotalsRow: shape is TableShape.Totals
        );
        return workbook;
    }

    public enum TableShape
    {
        Data,
        Totals,
        HeaderOnly,
    }
}
