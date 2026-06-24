using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using MemoryPack;

namespace Danfma.MySheet.Tests.Parsing;

public class FunctionExtensionTests
{
    private static (Workbook Workbook, Sheet Sheet) NewSheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add("Sheet1");

        return (workbook, sheet);
    }

    [Test]
    public async Task CustomFunction_IsResolvedAndInvoked()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction(
            "DOUBLE",
            (args, wb) => (args[0].Compute(wb) as double? ?? 0) * 2
        );

        await Assert
            .That(ExpressionParser.Parse("=DOUBLE(21)", sheet).Compute(workbook) as double?)
            .IsEqualTo(42.0);
    }

    [Test]
    public async Task UnregisteredFunction_IsNameError()
    {
        var (workbook, sheet) = NewSheet();

        await Assert
            .That(ExpressionParser.Parse("=NOPE()", sheet).Compute(workbook))
            .IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task XlfnPrefix_IsNormalized()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction("MYFN", (_, _) => 7.0);

        await Assert
            .That(ExpressionParser.Parse("=XLFN.MYFN()", sheet).Compute(workbook) as double?)
            .IsEqualTo(7.0);
        await Assert
            .That(ExpressionParser.Parse("=_xlfn.MYFN()", sheet).Compute(workbook) as double?)
            .IsEqualTo(7.0);
    }

    [Test]
    public async Task UnderscoreInFunctionName_Tokenizes()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction("A_HIDE", (_, _) => true);

        await Assert
            .That(ExpressionParser.Parse("=A_HIDE()", sheet).Compute(workbook) as bool?)
            .IsTrue();
    }

    [Test]
    public async Task CustomFunctionCall_SurvivesSerialization()
    {
        var (workbook, sheet) = NewSheet();
        sheet["A1"] = ExpressionParser.Parse("=CUSTOM(1,2)", sheet);

        var bytes = MemoryPackSerializer.Serialize(workbook);
        var restored = MemoryPackSerializer.Deserialize<Workbook>(bytes)!;

        // The FunctionCall node round-trips (name + args); the registry is not persisted, so re-register.
        restored.RegisterFunction(
            "CUSTOM",
            (args, wb) =>
                (args[0].Compute(wb) as double? ?? 0) + (args[1].Compute(wb) as double? ?? 0)
        );

        await Assert.That(restored["Sheet1"]["A1"].Compute(restored) as double?).IsEqualTo(3.0);
    }
}
