using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ArrayConstantTests
{
    private static object? Calc(string formula)
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        return ExpressionParser.Parse(formula, sheet).Evaluate(workbook).AsObject();
    }

    [Test]
    [Arguments("=SUM({1,2,3})", 6d)]
    [Arguments("=SUM({1;2;3})", 6d)]
    [Arguments("=SUM({1,2;3,4})", 10d)]
    [Arguments("=ROWS({1,2;3,4})", 2d)]
    [Arguments("=COLUMNS({1,2,3})", 3d)]
    [Arguments("=INDEX({1,2,3},2)", 2d)]
    [Arguments("=INDEX({1,2;3,4},2,1)", 3d)]
    [Arguments("=MATCH(2,{1,2,3},0)", 2d)]
    [Arguments("=XLOOKUP(2,{1,2,3},{10,20,30})", 20d)]
    [Arguments("=VLOOKUP(2,{1,10;2,20},2,FALSE)", 20d)]
    [Arguments("=SUMPRODUCT({1,2},{3,4})", 11d)]
    [Arguments("=SUM({1,2,3}*2)", 12d)]
    [Arguments("=SUM(IF({TRUE,FALSE},1,2))", 3d)]
    [Arguments("=LET(a,{1,2,3},SUM(a))", 6d)]
    public async Task ArrayConstants_WorkAcrossExistingArrayConsumers(
        string formula,
        double expected
    ) => await Assert.That(Calc(formula) as double?).IsEqualTo(expected);

    [Test]
    public async Task ArrayConstants_PreserveLiteralElementKinds()
    {
        await Assert.That(Calc("=INDEX({1,\"a\",TRUE,#N/A},1,2)") as string).IsEqualTo("a");
        await Assert.That(Calc("=INDEX({1,\"a\",TRUE,#N/A},1,3)") as bool?).IsTrue();
        await Assert
            .That(Calc("=INDEX({1,\"a\",TRUE,#N/A},1,4)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=COUNTA({1,\"a\",TRUE,#N/A})") as double?).IsEqualTo(4d);
        await Assert.That(Calc("=COUNT({1,\"a\",TRUE,#N/A})") as double?).IsEqualTo(1d);
        await Assert.That(Calc("=SUM({-1,2,1.5E3})") as double?).IsEqualTo(1501d);
    }

    [Test]
    public async Task BareArrayConstant_ReturnsItsTopLeftElement() =>
        await Assert.That(Calc("={1,2,3}") as double?).IsEqualTo(1d);

    [Test]
    public async Task CriteriaRangeSlot_RefusesArrayConstant() =>
        await Assert.That(Calc("=COUNTIF({1,2,3},\">1\")")).IsEqualTo(ErrorValue.Reference);

    [Test]
    public async Task LogicalAndTextConsumers_ReadArrayConstant()
    {
        await Assert.That(Calc("=AND({TRUE,FALSE})") as bool?).IsFalse();
        await Assert.That(Calc("=TEXTJOIN(\",\",TRUE,{\"a\",\"b\"})") as string).IsEqualTo("a,b");
        await Assert.That(Calc("=INDEX({1,2,3},-1)")).IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task DefinedName_CanBindArrayConstant()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        workbook.DefineName("Items", ExpressionParser.Parse("={1,2,3}", sheet));

        await Assert
            .That(ExpressionParser.Parse("=SUM(Items)", sheet).Evaluate(workbook).ToDouble())
            .IsEqualTo(6d);
    }

    [Test]
    public async Task OneByOneLookupArray_UsesTheFirstReturnValue()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["B1"] = new NumberValue(10);
        sheet["B2"] = new NumberValue(20);
        sheet["B3"] = new NumberValue(30);

        await Assert
            .That(
                ExpressionParser.Parse("=XLOOKUP(1,{1},B1:B3)", sheet).Evaluate(workbook).ToDouble()
            )
            .IsEqualTo(10d);
    }

    [Test]
    [Arguments("={1,2;3}")]
    [Arguments("={1+1}")]
    [Arguments("={A1}")]
    [Arguments("={1,,3}")]
    [Arguments("={1,2,}")]
    [Arguments("={1,+2}")]
    [Arguments("={-0}")]
    [Arguments("={1E+3,.5,-0,1e-3}")]
    [Arguments("={-0.0}")]
    [Arguments("={-0E+5}")]
    [Arguments("={-.0}")]
    public async Task InvalidArrayConstants_AreRejectedAtEntry(string formula) =>
        await Assert
            .That(() => ExpressionParser.Parse(formula, new Sheet { Name = "Sheet1" }))
            .Throws<ParseException>();

    [Test]
    [Arguments("={-1}", -1d)]
    [Arguments("={0}", 0d)]
    public async Task ValidSignedArrayNumbers_AreAcceptedAtEntry(string formula, double expected) =>
        await Assert.That(Calc(formula) as double?).IsEqualTo(expected);

    [Test]
    [Arguments("=ROW({1})")]
    [Arguments("=COLUMN({1})")]
    [Arguments("=ROW(SEQUENCE(2))")]
    [Arguments("=COLUMN(SEQUENCE(1,2))")]
    public async Task RowAndColumn_RejectComputedArrays(string formula) =>
        await Assert.That(Calc(formula)).IsEqualTo(ErrorValue.Reference);

    [Test]
    public async Task Rows_StillReadsAnArrayConstantShape() =>
        await Assert.That(Calc("=ROWS({1,2})") as double?).IsEqualTo(1d);

    [Test]
    public async Task Row_ScalarUsesTheNonReferenceFallback() =>
        // This was #VALUE! before the Aspose.Cells 26.7.0 measurement established #REF! in both modes.
        await Assert.That(Calc("=ROW(1)")).IsEqualTo(ErrorValue.Reference);

    [Test]
    public async Task LogicalConsumers_UseTheSharedArrayProducerRoute() =>
        await Assert.That(Calc("=AND(SEQUENCE(1,2)>1)") as bool?).IsFalse();

    [Test]
    public async Task LookupConsumers_UseTheSharedArrayProducerRoute()
    {
        await Assert
            .That(Calc("=VLOOKUP(2,SEQUENCE(2,2),2,FALSE)"))
            .IsEqualTo(ErrorValue.NotAvailable);
        await Assert.That(Calc("=HLOOKUP(2,SEQUENCE(2,2),2,FALSE)") as double?).IsEqualTo(4d);
        await Assert.That(Calc("=MATCH(2,SEQUENCE(3),0)") as double?).IsEqualTo(2d);
    }

    [Test]
    public async Task FormulaTextAndMemoryPack_RoundTripArrayConstant()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = ExpressionParser.Parse("={ 1 , \"a\" ; TRUE , #N/A }", sheet);
        sheet["B1"] = ExpressionParser.Parse("=FORMULATEXT(A1)", sheet);
        var path = Path.Combine(Path.GetTempPath(), $"mysheet-array-{Guid.NewGuid():N}.bin");

        try
        {
            await Assert.That(sheet["A1"].ToFormula("Sheet1")).IsEqualTo("{1,\"a\";TRUE,#N/A}");
            await Assert
                .That(workbook.GetCellValue("Sheet1", "B1").AsString())
                .IsEqualTo("={1,\"a\";TRUE,#N/A}");

            workbook.Save(path);
            var loaded = Workbook.Load(path);
            await Assert.That(loaded["Sheet1"]["A1"].GetType().Name).IsEqualTo("ArrayConstant");
            await Assert.That(loaded.GetCellValue("Sheet1", "A1").ToDouble()).IsEqualTo(1d);
            await Assert
                .That(loaded["Sheet1"]["A1"].ToFormula("Sheet1"))
                .IsEqualTo("{1,\"a\";TRUE,#N/A}");
        }
        finally
        {
            File.Delete(path);
        }
    }
}
