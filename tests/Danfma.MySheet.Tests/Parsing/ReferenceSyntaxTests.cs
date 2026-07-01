using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Parsing;

public class ReferenceSyntaxTests
{
    [Test]
    public async Task SheetNames_AreCaseInsensitive()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        var other = workbook.Sheets.Add("Other");
        sheet1["A1"] = new NumberValue(5);

        await Assert
            .That(ExpressionParser.Parse("=SHEET1!A1", other).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task SheetNames_AreCaseInsensitive_AfterRoundTrip()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        sheet1["A1"] = new NumberValue(5);

        var restored = MemoryPackSerializer.Deserialize<Workbook>(
            MemoryPackSerializer.Serialize(workbook)
        )!;
        var context = restored.Sheets.Add("Other");

        await Assert
            .That(ExpressionParser.Parse("=sheet1!A1", context).Evaluate(restored).AsObject() as double?)
            .IsEqualTo(5.0);
    }

    [Test]
    public async Task AbsoluteMarkers_AreIgnored()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(7);

        await Assert
            .That(ExpressionParser.Parse("=$A$1", sheet).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(7.0);
        await Assert
            .That(ExpressionParser.Parse("=$A1+A$1", sheet).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(14.0);
    }

    [Test]
    public async Task SheetQualifiedCell()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        var sheet2 = workbook.Sheets.Add("Sheet2");
        sheet2["A1"] = new NumberValue(42);

        await Assert
            .That(ExpressionParser.Parse("=Sheet2!A1", sheet1).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(42.0);
        await Assert
            .That(ExpressionParser.Parse("=Sheet2!$A$1", sheet1).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(42.0);
    }

    [Test]
    public async Task SheetQualifiedRange()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        var sheet2 = workbook.Sheets.Add("Sheet2");
        sheet2["A1"] = new NumberValue(1);
        sheet2["A2"] = new NumberValue(2);
        sheet2["A3"] = new NumberValue(3);

        await Assert
            .That(ExpressionParser.Parse("=SUM(Sheet2!A1:A3)", sheet1).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(6.0);
    }

    [Test]
    public async Task QuotedSheetName()
    {
        var workbook = new Workbook();
        var sheet1 = workbook.Sheets.Add("Sheet1");
        var other = workbook.Sheets.Add("My Sheet");
        other["B2"] = new NumberValue(99);

        await Assert
            .That(ExpressionParser.Parse("='My Sheet'!B2", sheet1).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(99.0);
    }

    [Test]
    public async Task UnionOperator_CombinesAreas()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");
        sheet["A1"] = new NumberValue(1);
        sheet["A2"] = new NumberValue(2);
        sheet["A3"] = new NumberValue(3);
        sheet["C1"] = new NumberValue(10);
        sheet["C2"] = new NumberValue(20);
        sheet["C3"] = new NumberValue(30);

        await Assert
            .That(ExpressionParser.Parse("=SUM((A1:A3,C1:C3))", sheet).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(66.0);
        await Assert
            .That(ExpressionParser.Parse("=COUNT((A1,A3,C2))", sheet).Evaluate(workbook).AsObject() as double?)
            .IsEqualTo(3.0);
    }
}
