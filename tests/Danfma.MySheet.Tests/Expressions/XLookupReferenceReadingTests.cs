using System.Globalization;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using StringValue = Danfma.MySheet.Expressions.StringValue;

namespace Danfma.MySheet.Tests.Expressions;

public class XLookupReferenceReadingTests
{
    // Fixture: Main!A1:A3=1,2,3; B1:B3=10,20,30; C1:C3=100,200,300; formula at AZ5000.
    // Aspose.Cells 26.7.0 PLAIN/CSE agree. Before this fix, reference readers rejected direct and
    // computed-reference returns (#REF! for criteria, #VALUE! for aggregates) or collapsed them to 1/20.
    [Test]
    [Arguments("=COUNTIF(XLOOKUP(2,A1:A3,B1:C3),\">0\")", "2")]
    [Arguments("=SUMIF(XLOOKUP(2,A1:A3,B1:C3),\">0\")", "220")]
    [Arguments("=AGGREGATE(9,6,XLOOKUP(2,A1:A3,B1:C3))", "220")]
    [Arguments("=SUBTOTAL(9,XLOOKUP(2,A1:A3,B1:C3))", "220")]
    [Arguments("=COUNTIFS(XLOOKUP(2,A1:A3,B1:C3),\">0\")", "2")]
    [Arguments("=AVERAGEIF(XLOOKUP(2,A1:A3,B1:C3),\">0\")", "110")]
    [Arguments("=MAXIFS(XLOOKUP(2,A1:A3,B1:C3),XLOOKUP(2,A1:A3,B1:C3),\">0\")", "200")]
    [Arguments("=COUNTIF(XLOOKUP(2,SEQUENCE(3),B1:C3),\">0\")", "2")]
    [Arguments("=COUNTIF(XLOOKUP(2,A:A,B:C),\">0\")", "2")]
    [Arguments("=SUM(OFFSET(XLOOKUP(2,A1:A3,B1:C3),0,0))", "220")]
    [Arguments("=COUNTIF(XLOOKUP(2,SEQUENCE(3),SEQUENCE(3,3)),\">0\")", "#REF!")]
    [Arguments("=AGGREGATE(9,6,XLOOKUP(2,SEQUENCE(3),SEQUENCE(3,3)))", "#VALUE!")]
    [Arguments("=LET(r,B1:C3,COUNTIF(XLOOKUP(2,A1:A3,r),\">0\"))", "#REF!")]
    public async Task ReferenceReaders_UseTheSelectedXLookupReference(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    // Same fixture and oracle protocol. The old MySheet values were #REF! for every criteria row except
    // the single-cell shape (#VALUE!); the accepted rows now read the selected reference in full.
    [Test]
    [Arguments("B1:C3", "2")]
    [Arguments("B2", "#VALUE!")]
    [Arguments("RetRng", "2")]
    [Arguments("Tabela1[Valor]", "1")]
    [Arguments("Other!B1:C3", "2")]
    [Arguments("INDEX(B1:C3,0,0)", "2")]
    [Arguments("OFFSET(B1,0,0,3,2)", "2")]
    [Arguments("IF(TRUE,B1:C3,B1:C3)", "2")]
    [Arguments("CHOOSE(1,B1:C3)", "2")]
    [Arguments("INDIRECT(\"B1:C3\")", "2")]
    public async Task ReturnOperandSyntax_DecidesWhetherXLookupReturnsAReference(
        string returnOperand,
        string expected
    ) =>
        await Assert
            .That(Evaluate($"=COUNTIF(XLOOKUP(2,A1:A3,{returnOperand}),\">0\")"))
            .IsEqualTo(expected);

    [Test]
    [Arguments("=COUNTIF(XLOOKUP(2,A1:A3,A1:A3),\">0\")", "#N/A")]
    [Arguments("=SUMIF(XLOOKUP(2,A1:A3,A1:A3),\">0\")", "#N/A")]
    [Arguments("=COUNTIF(XLOOKUP(2,A1:A3,A1:A3,B1:B3),\">0\")", "#VALUE!")]
    public async Task LookupMiss_PropagatesThroughReferenceReaders(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula, missFixture: true)).IsEqualTo(expected);

    // Aspose.Cells 26.7.0 PLAIN/CSE agree. Before this fix, reference-only readers re-evaluated
    // scalar fallbacks as ranges: COUNTIF returned 0/1 and SUMIF returned 0 instead of the slot error.
    [Test]
    [Arguments("=COUNTIF(XLOOKUP(9,A1:A3,B1:C3,0),\">0\")", "#REF!")]
    [Arguments("=COUNTIF(XLOOKUP(9,A1:A3,B1:C3,E1),\">0\")", "#REF!")]
    [Arguments("=SUMIF(XLOOKUP(9,A1:A3,B1:C3,0),\">0\")", "#REF!")]
    [Arguments("=AGGREGATE(9,6,XLOOKUP(9,A1:A3,B1:C3,0))", "#VALUE!")]
    [Arguments("=SUBTOTAL(9,XLOOKUP(9,A1:A3,B1:C3,0))", "#VALUE!")]
    [Arguments("=SUM(OFFSET(XLOOKUP(9,A1:A3,B1:C3,0),0,0))", "#REF!")]
    [Arguments("=COUNTIF(XLOOKUP(9,A1:A3,B1:C3,\"nf\"),\">0\")", "#REF!")]
    [Arguments("=COUNTIF(XLOOKUP(9,A1:A3,B1:C3,NA()),\">0\")", "#N/A")]
    public async Task ScalarFallback_InAReferenceOnlySlot_ReportsTheMeasuredSlotError(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    [Test]
    [Arguments("=COUNTIF(LET(r,XLOOKUP(2,A1:A3,B1:C3),r),\">0\")", "#REF!")]
    [Arguments("=COUNTIF(LET(r,XLOOKUP(2,A1:A3,B1:C3),s,r,s),\">0\")", "#REF!")]
    [Arguments("=SUMIF(A1:A3,\">0\",LET(r,XLOOKUP(2,A1:A3,B1:C3),r))", "#REF!")]
    [Arguments("=AGGREGATE(9,6,LET(r,XLOOKUP(2,A1:A3,B1:C3),r))", "#VALUE!")]
    [Arguments("=SUM(OFFSET(LET(r,XLOOKUP(2,A1:A3,B1:C3),r),0,0))", "#REF!")]
    [Arguments("=SUM(LET(r,XLOOKUP(2,A1:A3,B1:C3),r))", "220")]
    public async Task LetBoundXLookup_IsAValueArrayRatherThanAStructuralReference(
        string formula,
        string expected
    )
    {
        // Aspose.Cells 26.7.0 PLAIN/CSE agree on every row. MySheet previously promoted the LET name
        // to a reference, producing 2, 2, 50, 20, #REF!, and 20 respectively instead of these values.
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=COUNTIF(LET(r,INDEX(A1:B3,0,1),r),\">0\")", "3")]
    [Arguments("=COUNTIF(XLOOKUP(2,A1:A3,B1:C3),\">0\")", "2")]
    [Arguments("=COUNTIF(LET(r,Ghost!A1:A3,r),\">0\")", "#REF!")]
    public async Task LetBoundXLookup_RetainsReferenceRouteControls(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    [Test]
    [Arguments("=SUM(OFFSET(XLOOKUP(2,A1:A3,B1:C3),0,0,-1,1))", "20")]
    [Arguments("=SUM(OFFSET(XLOOKUP(3,A1:A3,B1:C3),0,0,-2,2))", "550")]
    public async Task SelectedReference_PreservesNegativeOffsetGeometry(
        string formula,
        string expected
    ) => await Assert.That(Evaluate(formula)).IsEqualTo(expected);

    [Test]
    [Arguments("=COUNTIF(XLOOKUP(TICK(),A1:A3,B1:C3),\">0\")", "2")]
    [Arguments("=AGGREGATE(9,6,XLOOKUP(TICK(),A1:A3,B1:C3))", "220")]
    public async Task ReferenceReaders_EvaluateVolatileXLookupOnce(string formula, string expected)
    {
        var workbook = CreateWorkbook();
        var draws = 0;
        workbook.RegisterFunction(
            "TICK",
            (_, _) =>
            {
                draws++;
                return 2;
            }
        );

        await Assert.That(Evaluate(workbook, formula)).IsEqualTo(expected);
        await Assert.That(draws).IsEqualTo(1);
    }

    private static string Evaluate(string formula, bool missFixture = false)
    {
        var workbook = CreateWorkbook(missFixture);
        return Evaluate(workbook, formula);
    }

    private static Workbook CreateWorkbook(bool missFixture = false)
    {
        var workbook = new Workbook();
        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(missFixture ? 5 : 1);
        main["A2"] = new NumberValue(missFixture ? 0 : 2);
        main["A3"] = new NumberValue(missFixture ? 9 : 3);
        main["B1"] = new NumberValue(10);
        main["B2"] = new NumberValue(20);
        main["B3"] = new NumberValue(30);
        main["C1"] = new NumberValue(100);
        main["C2"] = new NumberValue(200);
        main["C3"] = new NumberValue(300);

        var other = workbook.Sheets.Add("Other");
        other["B1"] = new NumberValue(10);
        other["B2"] = new NumberValue(20);
        other["B3"] = new NumberValue(30);
        other["C1"] = new NumberValue(100);
        other["C2"] = new NumberValue(200);
        other["C3"] = new NumberValue(300);

        var data = workbook.Sheets.Add("Data");
        data["A1"] = new StringValue("Valor");
        data["A2"] = new NumberValue(10);
        data["A3"] = new NumberValue(20);
        data["A4"] = new NumberValue(30);
        workbook.DefineTable("Tabela1", "Data", "A1:A4", ["Valor"]);
        workbook.DefineName("RetRng", "Main!$B$1:$C$3");
        return workbook;
    }

    private static string Evaluate(Workbook workbook, string formula)
    {
        var main = workbook.Sheets["Main"];
        main["AZ5000"] = ExpressionParser.Parse(formula, main);
        var value = workbook.GetCellValue("Main", "AZ5000");

        return value.TryGetError(out var error) ? error.ToString()
            : value.TryGetNumber(out var number) ? number.ToString(CultureInfo.InvariantCulture)
            : value.Kind.ToString();
    }
}
