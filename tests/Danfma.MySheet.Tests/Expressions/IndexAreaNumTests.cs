using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Expressions;

/// <summary>
/// INDEX area_num oracle: Aspose.Cells 26.7.0, one formula per workbook at Main!AZ5000,
/// measured 2026-09-15. PLAIN and CSE agreed on every row below.
/// </summary>
public class IndexAreaNumTests
{
    [Test]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,1)", 2.0)]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,2)", 20.0)]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,1.9)", 2.0)]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,\"2\")", 20.0)]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1)", 2.0)]
    [Arguments("=INDEX((A1:A3,C1:C3,E1:E2),2,1,3)", 300.0)]
    public async Task AreaNum_SelectsAOneBasedAreaBeforeApplyingAxes(
        string formula,
        double expected
    )
    {
        await Assert.That(Eval(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,3)", "#REF!")]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,0)", "#VALUE!")]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,-1)", "#VALUE!")]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,1/0)", "#DIV/0!")]
    [Arguments("=INDEX((A1:A3,C1:C3),-1,1,2)", "#VALUE!")]
    public async Task AreaNum_UsesTheOracleErrorPrecedence(string formula, string expected)
    {
        await Assert.That(((ErrorValue)Eval(formula)!).ErrorCode).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=INDEX(A1:A3,2,1,1)", 2.0)]
    [Arguments("=INDEX(A1,1,1,1)", 1.0)]
    public async Task AreaNumOne_AcceptsASingleReferenceArea(string formula, double expected)
    {
        await Assert.That(Eval(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=INDEX(A1:A3,2,1,2)")]
    [Arguments("=INDEX(A1,1,1,2)")]
    public async Task AreaNumPastASingleReferenceArea_IsRefError(string formula)
    {
        await Assert.That(Eval(formula)).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    public async Task AreaNum_RejectsANonReferenceArray()
    {
        await Assert.That(Eval("=INDEX(SEQUENCE(3),2,1,1)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    [Arguments("=SUM(INDEX((A1:A3,C1:C3),0,1,2))", 60.0)]
    [Arguments("=ROWS(INDEX((A1:A3,E1:F2),0,0,2))", 2.0)]
    [Arguments("=AREAS(INDEX((A1:A3,C1:C3),0,0,2))", 1.0)]
    public async Task SelectedArea_RemainsAReferenceForReferenceAwareConsumers(
        string formula,
        double expected
    )
    {
        await Assert.That(Eval(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=INDEX((Ghost!A1:A3,C1:C3),2,1,2)", 20.0)]
    [Arguments("=SUM(INDEX((Ghost!A1:A3,C1:C3),0,1,2))", 60.0)]
    public async Task MissingSheet_InAnUnselectedArea_DoesNotPoisonTheSelectedArea(
        string formula,
        double expected
    )
    {
        await Assert.That(Eval(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("=INDEX((A1:A3,Ghost!C1:C3),2,1,2)")]
    [Arguments("=SUM(INDEX((A1:A3,Ghost!C1:C3),0,1,2))")]
    public async Task MissingSheet_InTheSelectedArea_IsRefError(string formula)
    {
        await Assert.That(Eval(formula)).IsEqualTo(ErrorValue.Reference);
    }

    [Test]
    [Arguments("=INDEX((A:A,C:C),2,1,2)", 20.0)]
    [Arguments("=INDEX((Main!A1:A3,Other!A1:A3),2,1,2)", 2000.0)]
    [Arguments("=OFFSET(INDEX((A1:A3,C1:C3),1,1,2),1,0)", 20.0)]
    [Arguments("=SUM(INDEX((A1:A3,C1:C3),0,0,2):C3)", 60.0)]
    [Arguments("=LET(u,(A1:A3,C1:C3),INDEX(u,2,1,2))", 20.0)]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,D1)", 20.0)]
    public async Task AreaNum_AdversarialReferenceShapesMatchTheOracle(
        string formula,
        double expected
    )
    {
        await Assert.That(Eval(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task AreaNum_FromABlankCell_IsValueError()
    {
        await Assert.That(Eval("=INDEX((A1:A3,C1:C3),2,1,D2)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task DefinedNameUnion_UsesTheCoherentAreaSemantics()
    {
        var workbook = Fixture();
        workbook.DefineName("U", "(Main!$A$1:$A$3,Main!$C$1:$C$3)");

        // Aspose 26.7.0 answers #VALUE! in PLAIN and CSE despite AREAS(U)=2 and the equivalent direct and
        // LET-bound unions returning 20. This pin keeps the coherent selected-area result as an oracle defect.
        await Assert.That(Eval(workbook, "=INDEX(U,2,1,2)")).IsEqualTo(20.0);
    }

    [Test]
    [Arguments("=INDEX((A1:A3,C1:C3),2,1,TICK())", 20.0)]
    [Arguments("=SUM(INDEX((A1:A3,C1:C3),0,1,TICK()))", 60.0)]
    [Arguments("=ROWS(INDEX((A1:A3,C1:C3),0,1,TICK()))", 3.0)]
    public async Task AreaNum_IsEvaluatedOnce(string formula, double expected)
    {
        var workbook = Fixture();
        var draws = 0;
        workbook.RegisterFunction(
            "TICK",
            (_, _) =>
            {
                draws++;
                return 2;
            }
        );

        await Assert.That(Eval(workbook, formula)).IsEqualTo(expected);
        await Assert.That(draws).IsEqualTo(1);
    }

    [Test]
    public async Task AreaNum_IsEvaluatedOnceWhenResolvingAReferenceDirectly()
    {
        var workbook = Fixture();
        var draws = 0;
        workbook.RegisterFunction(
            "TICK",
            (_, _) =>
            {
                draws++;
                return 2;
            }
        );
        var expression = ExpressionParser.Parse(
            "=INDEX((A1:A3,C1:C3),2,1,TICK())",
            workbook["Main"]
        );

        await Assert
            .That(
                expression.TryResolveReference(new EvaluationContext(workbook), out var reference)
            )
            .IsTrue();
        await Assert.That(reference).IsEqualTo(new CellReference("C2", "Main"));
        await Assert.That(draws).IsEqualTo(1);
    }

    [Test]
    public async Task AreaNum_CanSelectATableColumnFromAUnion()
    {
        var workbook = Fixture();
        workbook.DefineTable("Table1", "Main", "G1:H4", ["Value", "Other"]);

        await Assert.That(Eval(workbook, "=INDEX((Table1[Value],C1:C3),2,1,1)")).IsEqualTo(8.0);
    }

    private static object? Eval(string formula) => Eval(Fixture(), formula);

    private static object? Eval(Workbook workbook, string formula)
    {
        var sheet = workbook["Main"];
        sheet["AZ5000"] = ExpressionParser.Parse(formula, sheet);
        return workbook.GetCellValue("Main", "AZ5000").AsObject();
    }

    private static Workbook Fixture()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Main");
        var other = workbook.Sheets.Add("Other");

        foreach (
            var (cell, value) in new (string, double)[]
            {
                ("A1", 1),
                ("A2", 2),
                ("A3", 3),
                ("C1", 10),
                ("C2", 20),
                ("C3", 30),
                ("E1", 100),
                ("F1", 200),
                ("E2", 300),
                ("F2", 400),
                ("G2", 7),
                ("G3", 8),
                ("G4", 9),
            }
        )
        {
            sheet[cell] = new NumberValue(value);
        }

        sheet["G1"] = new Danfma.MySheet.Expressions.StringValue("Value");
        sheet["H1"] = new Danfma.MySheet.Expressions.StringValue("Other");
        sheet["D1"] = new Danfma.MySheet.Expressions.StringValue("2");
        other["A1"] = new NumberValue(1000);
        other["A2"] = new NumberValue(2000);
        other["A3"] = new NumberValue(3000);
        return workbook;
    }
}
