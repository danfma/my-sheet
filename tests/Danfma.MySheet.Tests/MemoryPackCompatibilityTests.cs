using System.Runtime.CompilerServices;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;
using Excel.FinancialFunctions;

namespace Danfma.MySheet.Tests;

/// <summary>
/// MemoryPack binary-compatibility guard. <c>Fixtures/workbook-pre-namespaces.msgpack.bin</c> was
/// serialized by <see cref="Workbook.Save"/> BEFORE the 2.0 namespace reorganization, from the workbook
/// built by <see cref="BuildRepresentativeWorkbook"/> — formulas from every function category
/// (arithmetic operators, SUM/AVERAGE over ranges, IF/IFS, text, VLOOKUP/CHOOSE, PMT, a 1/0 error,
/// cross-sheet references, a reference union, LET and a custom <c>FunctionCall</c>). The permanent test
/// proves the tag-based <c>[MemoryPackUnion]</c> contract survives reorganizing the AST nodes into new
/// namespaces: the frozen file must keep loading and re-evaluating identically, whatever the type layout.
/// </summary>
public class MemoryPackCompatibilityTests
{
    private const string FixtureFileName = "workbook-pre-namespaces.msgpack.bin";
    private const double Tolerance = 1e-6;

    [Test]
    public async Task PreNamespaceFixture_LoadsAndReevaluates()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", FixtureFileName);
        var workbook = Workbook.Load(path);

        // The custom-function registry is not serialized; hosts re-register after loading.
        workbook.RegisterFunction(
            "DOUBLE",
            (args, wb) => (args[0].Evaluate(wb).AsObject() as double? ?? 0) * 2
        );

        object? Value(string id) => workbook["Main"][id].Evaluate(workbook).AsObject();

        // Operators: =A1+A2*2-3^2 (binary +, *, -, ^ and a literal).
        await Assert.That(Value("B1") as double?).IsEqualTo(41.0);

        // Mathematics over a range: =SUM(A1:A3).
        await Assert.That(Value("B2") as double?).IsEqualTo(60.0);

        // Statistical over a range: =AVERAGE(A1:A3).
        await Assert.That(Value("B3") as double?).IsEqualTo(20.0);

        // Logical: =IF(A1>5, "big", "small") and =IFS(A1>100, 1, A1>5, 2).
        await Assert.That(Value("B4") as string).IsEqualTo("big");
        await Assert.That(Value("B5") as double?).IsEqualTo(2.0);

        // Text: =UPPER(CONCAT("my", "sheet")).
        await Assert.That(Value("B6") as string).IsEqualTo("MYSHEET");

        // Lookup, cross-sheet table: =VLOOKUP("banana", Data!D1:E3, 2, FALSE) and =CHOOSE(3, ...).
        await Assert.That(Value("B7") as double?).IsEqualTo(2.0);
        await Assert.That(Value("B8") as string).IsEqualTo("c");

        // Financial: =PMT(0.05/12, 360, 200000), cross-checked against the ExcelFinancialFunctions oracle.
        var expectedPmt = Financial.Pmt(0.05 / 12, 360, 200000, 0, PaymentDue.EndOfPeriod);
        await Assert.That(Value("B9") as double? ?? double.NaN).IsEqualTo(expectedPmt).Within(Tolerance);

        // Error: =1/0.
        await Assert.That(Value("B10")).IsEqualTo(ErrorValue.DivByZero);

        // Cross-sheet cell reference: =Data!A1*3.
        await Assert.That(Value("B11") as double?).IsEqualTo(30.0);

        // Reference union: =SUM((A1:A2,A3:A3)).
        await Assert.That(Value("B12") as double?).IsEqualTo(60.0);

        // LET (name binding + NameReference): =LET(x, A1*2, x+x).
        await Assert.That(Value("B13") as double?).IsEqualTo(40.0);

        // Custom function call (extensibility node): =DOUBLE(A2).
        await Assert.That(Value("B14") as double?).IsEqualTo(40.0);
    }

    /// <summary>One workbook exercising every function category and reference shape we serialize.</summary>
    private static Workbook BuildRepresentativeWorkbook()
    {
        var workbook = new Workbook();

        var data = workbook.Sheets.Add("Data");
        data["A1"] = new NumberValue(10);
        data["D1"] = Expression.String("apple");
        data["D2"] = Expression.String("banana");
        data["D3"] = Expression.String("cherry");
        data["E1"] = new NumberValue(1);
        data["E2"] = new NumberValue(2);
        data["E3"] = new NumberValue(3);

        var main = workbook.Sheets.Add("Main");
        main["A1"] = new NumberValue(10);
        main["A2"] = new NumberValue(20);
        main["A3"] = new NumberValue(30);
        main["B1"] = ExpressionParser.Parse("=A1+A2*2-3^2", main);
        main["B2"] = ExpressionParser.Parse("=SUM(A1:A3)", main);
        main["B3"] = ExpressionParser.Parse("=AVERAGE(A1:A3)", main);
        main["B4"] = ExpressionParser.Parse("=IF(A1>5, \"big\", \"small\")", main);
        main["B5"] = ExpressionParser.Parse("=IFS(A1>100, 1, A1>5, 2)", main);
        main["B6"] = ExpressionParser.Parse("=UPPER(CONCAT(\"my\", \"sheet\"))", main);
        main["B7"] = ExpressionParser.Parse("=VLOOKUP(\"banana\", Data!D1:E3, 2, FALSE)", main);
        main["B8"] = ExpressionParser.Parse("=CHOOSE(3, \"a\", \"b\", \"c\")", main);
        main["B9"] = ExpressionParser.Parse("=PMT(0.05/12, 360, 200000)", main);
        main["B10"] = ExpressionParser.Parse("=1/0", main);
        main["B11"] = ExpressionParser.Parse("=Data!A1*3", main);
        main["B12"] = ExpressionParser.Parse("=SUM((A1:A2,A3:A3))", main);
        main["B13"] = ExpressionParser.Parse("=LET(x, A1*2, x+x)", main);
        main["B14"] = ExpressionParser.Parse("=DOUBLE(A2)", main);

        return workbook;
    }

    // Fixture generator — intentionally NOT a [Test]. The fixture was produced ONCE by the pre-2.0
    // (pre-namespace-reorg) layout and is frozen in git; regenerating it with current code would defeat
    // the whole point of the guard (it would no longer prove backward compatibility). If you must extend
    // it deliberately (e.g. new categories in a future major), temporarily add an instance wrapper
    //   [Test] public void GenerateFixtureRun() => GenerateFixture();
    // (TUnit tests must be non-static and parameterless), run
    //   dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release \
    //     -- --treenode-filter "/*/*/MemoryPackCompatibilityTests/GenerateFixtureRun"
    // then remove the wrapper and commit the new file together with updated expectations above.
    private static void GenerateFixture([CallerFilePath] string sourcePath = "")
    {
        var directory = Path.Combine(Path.GetDirectoryName(sourcePath)!, "Fixtures");
        Directory.CreateDirectory(directory);
        BuildRepresentativeWorkbook().Save(Path.Combine(directory, FixtureFileName));
    }
}
