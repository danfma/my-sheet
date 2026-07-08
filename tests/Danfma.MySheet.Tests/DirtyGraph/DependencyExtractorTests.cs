using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 1 do spike do grafo de dependências: extração forward de deps do AST + classificação always-dirty.
public class DependencyExtractorTests
{
    private static DependencyScan Scan(string formula, Workbook? workbook = null)
    {
        var sheet = new Sheet { Name = "Sheet1" };
        return DependencyExtractor.Extract(ExpressionParser.Parse(formula, sheet), workbook);
    }

    private static CellDep Cell(int col, int row, string sheet = "Sheet1") => new(sheet, col, row);

    [Test]
    public async Task SimpleArithmetic_CollectsBothCells()
    {
        var scan = Scan("=A1+B2");

        await Assert.That(scan.Cells.Count).IsEqualTo(2);
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
        await Assert.That(scan.Cells.Contains(Cell(2, 2))).IsTrue(); // B2
        await Assert.That(scan.Ranges.Count).IsEqualTo(0);
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task BoundedRange_IsASingleRangeDep()
    {
        var scan = Scan("=SUM(A1:C3)");

        await Assert.That(scan.Cells.Count).IsEqualTo(0);
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 3, 1, 3));
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task WholeColumn_IsAnOpenRangeDep()
    {
        var scan = Scan("=SUM(A:A)");

        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 1, null, null));
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task Union_CollectsEachArea()
    {
        var scan = Scan("=SUM((A1:A2,C1))");

        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 1, 1, 2)); // A1:A2
        await Assert.That(scan.Cells.Contains(Cell(3, 1))).IsTrue(); // C1
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task Offset_IsAlwaysDirty_ButStillCollectsTheBaseCell()
    {
        var scan = Scan("=OFFSET(A1,1,0)");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1 (a base é uma dep real)
    }

    [Test]
    public async Task Indirect_IsAlwaysDirty()
    {
        var scan = Scan("=B1&INDIRECT(B2)");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(2, 1))).IsTrue(); // B1
        await Assert.That(scan.Cells.Contains(Cell(2, 2))).IsTrue(); // B2
    }

    [Test]
    public async Task Index_IsNotDirty_TheWholeRangeIsTheDependency()
    {
        // INDEX varre o range inteiro (row/col computados), então a dep é o range — enumerável, não dirty.
        var scan = Scan("=INDEX(A1:C10,2,3)");

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 1, 3, 1, 10));
    }

    [Test]
    public async Task If_CollectsConditionAndBothBranches()
    {
        // Super-aproximação: ambos os ramos são deps (qualquer um pode ser o resultado).
        var scan = Scan("=IF(A1>5,B1,C1)");

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
        await Assert.That(scan.Cells.Contains(Cell(2, 1))).IsTrue(); // B1
        await Assert.That(scan.Cells.Contains(Cell(3, 1))).IsTrue(); // C1
    }

    [Test]
    public async Task Now_IsAlwaysDirty_WithNoCellDeps()
    {
        var scan = Scan("=NOW()");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Count).IsEqualTo(0);
        await Assert.That(scan.Ranges.Count).IsEqualTo(0);
    }

    [Test]
    public async Task CrossSheetReference_CarriesTheSheetName()
    {
        var scan = Scan("=Sheet2!A1*3");

        await Assert.That(scan.Cells.Contains(Cell(1, 1, "Sheet2"))).IsTrue();
        await Assert.That(scan.AlwaysDirty).IsFalse();
    }

    [Test]
    public async Task DynamicRangeEndpoint_IsAlwaysDirty()
    {
        // A1:INDEX(B1:B9,3) — endpoint reference-returning → DynamicRange → always-dirty, mas coleta o que dá.
        var scan = Scan("=SUM(A1:INDEX(B1:B9,3))");

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
        await Assert.That(scan.Ranges).Contains(new RangeDep("Sheet1", 2, 2, 1, 9)); // B1:B9
    }

    [Test]
    public async Task CustomFunctionCall_IsAlwaysDirty()
    {
        var scan = Scan("=MYFUNC(A1)"); // não é built-in → FunctionCall

        await Assert.That(scan.AlwaysDirty).IsTrue();
        await Assert.That(scan.Cells.Contains(Cell(1, 1))).IsTrue(); // A1
    }

    [Test]
    public async Task DefinedName_ResolvesToItsReference()
    {
        var workbook = new Workbook();
        workbook.DefineName("Sales", "Data!A1:A3");

        var scan = Scan("=SUM(Sales)", workbook);

        await Assert.That(scan.AlwaysDirty).IsFalse();
        await Assert.That(scan.Ranges).Contains(new RangeDep("Data", 1, 1, 1, 3));
    }

    [Test]
    public async Task UnresolvableName_IsAlwaysDirty()
    {
        var scan = Scan("=SUM(Unknown)"); // sem workbook → nome irresolúvel

        await Assert.That(scan.AlwaysDirty).IsTrue();
    }
}
