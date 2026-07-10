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
            (args, wb) => (args[0].Evaluate(wb).AsObject() as double? ?? 0) * 2
        );

        await Assert
            .That(
                ExpressionParser.Parse("=DOUBLE(21)", sheet).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(42.0);
    }

    [Test]
    public async Task UnregisteredFunction_IsNameError()
    {
        var (workbook, sheet) = NewSheet();

        await Assert
            .That(ExpressionParser.Parse("=NOPE()", sheet).Evaluate(workbook).AsObject())
            .IsEqualTo(ErrorValue.Name);
    }

    [Test]
    public async Task XlfnPrefix_IsNormalized()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction("MYFN", (_, _) => 7.0);

        await Assert
            .That(
                ExpressionParser.Parse("=XLFN.MYFN()", sheet).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(7.0);
        await Assert
            .That(
                ExpressionParser.Parse("=_xlfn.MYFN()", sheet).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(7.0);
    }

    [Test]
    public async Task UnderscoreInFunctionName_Tokenizes()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction("A_HIDE", (_, _) => true);

        await Assert
            .That(ExpressionParser.Parse("=A_HIDE()", sheet).Evaluate(workbook).AsObject() as bool?)
            .IsTrue();
    }

    [Test]
    public async Task CustomFunction_WithDeclaredArity_TooFewArguments_IsValueError()
    {
        var (workbook, sheet) = NewSheet();

        // Would throw IndexOutOfRangeException on args[1] if invoked with a single argument -- the arity
        // guard must reject the call before the delegate runs.
        workbook.RegisterFunction(
            "ADDTWO",
            (args, wb) =>
                (args[0].Evaluate(wb).AsObject() as double? ?? 0)
                + (args[1].Evaluate(wb).AsObject() as double? ?? 0),
            minArgs: 2,
            maxArgs: 2
        );

        await Assert
            .That(ExpressionParser.Parse("=ADDTWO(1)", sheet).Evaluate(workbook).AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task CustomFunction_WithDeclaredArity_TooManyArguments_IsValueError()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction("ADDTWO", (args, wb) => 0.0, minArgs: 2, maxArgs: 2);

        await Assert
            .That(ExpressionParser.Parse("=ADDTWO(1,2,3)", sheet).Evaluate(workbook).AsObject())
            .IsEqualTo(ErrorValue.NotValue);
    }

    [Test]
    public async Task CustomFunction_WithoutDeclaredArity_PreservesLegacyBehavior()
    {
        var (workbook, sheet) = NewSheet();

        // No minArgs/maxArgs supplied: the pre-existing, unchecked behavior -- calling with too few
        // arguments still reaches into the delegate and throws, exactly as it did before arity validation
        // was added, instead of being caught upstream as #VALUE!.
        workbook.RegisterFunction(
            "ADDTWO",
            (args, wb) => args[0].Evaluate(wb).AsObject() as double? ?? 0
        );

        await Assert
            .That(() => ExpressionParser.Parse("=ADDTWO()", sheet).Evaluate(workbook))
            .Throws<IndexOutOfRangeException>();
    }

    [Test]
    public async Task CustomFunction_WithDeclaredArity_WithinRange_IsInvoked()
    {
        var (workbook, sheet) = NewSheet();
        workbook.RegisterFunction(
            "ADDTWO",
            (args, wb) =>
                (args[0].Evaluate(wb).AsObject() as double? ?? 0)
                + (args[1].Evaluate(wb).AsObject() as double? ?? 0),
            minArgs: 2,
            maxArgs: 2
        );

        await Assert
            .That(
                ExpressionParser.Parse("=ADDTWO(1,2)", sheet).Evaluate(workbook).AsObject()
                    as double?
            )
            .IsEqualTo(3.0);
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
                (args[0].Evaluate(wb).AsObject() as double? ?? 0)
                + (args[1].Evaluate(wb).AsObject() as double? ?? 0)
        );

        await Assert
            .That(restored["Sheet1"]["A1"].Evaluate(restored).AsObject() as double?)
            .IsEqualTo(3.0);
    }
}
