using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.Parsing;

public class ChainedReferenceTests
{
    [Test]
    public async Task ChainedReferences_WithConcatenation()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        sheet["A1"] = new Danfma.MySheet.Expressions.StringValue("Hello");
        sheet["B1"] = ExpressionParser.Parse("=A1 & \", World\"", sheet);
        sheet["C1"] = ExpressionParser.Parse("=B1 & \"!\"", sheet);

        // C1 -> B1 -> A1
        await Assert
            .That(sheet["C1"].Evaluate(workbook).AsObject() as string)
            .IsEqualTo("Hello, World!");
    }

    [Test]
    public async Task ChainedReferences_SharedDependencyComputedOnce()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        var calls = 0;
        workbook.RegisterFunction(
            "NAME",
            (_, _) =>
            {
                calls++;
                return "Ada";
            }
        );

        sheet["A1"] = ExpressionParser.Parse("=NAME()", sheet);
        sheet["B1"] = ExpressionParser.Parse("=A1 & \" Lovelace\"", sheet);
        sheet["C1"] = ExpressionParser.Parse("=B1 & \" — \" & A1", sheet); // references A1 again

        await Assert
            .That(sheet["C1"].Evaluate(workbook).AsObject() as string)
            .IsEqualTo("Ada Lovelace — Ada");
        // A1 is reached via B1 and directly, but memoized → NAME() runs once.
        await Assert.That(calls).IsEqualTo(1);
    }
}
