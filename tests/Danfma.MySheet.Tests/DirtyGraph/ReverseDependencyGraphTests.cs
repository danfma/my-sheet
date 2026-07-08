using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 2: grafo reverso + índice de contenção de ranges. O GetAllDependents é validado contra um ORÁCULO
// por força bruta (fecho transitivo re-derivado varrendo TODAS as fórmulas), em cenários controlados e num
// workbook randomizado.
public class ReverseDependencyGraphTests
{
    private static CellDep C(int col, int row, string sheet = "Sheet1") => new(sheet, col, row);

    private static CellDep Addr(string sheet, string id)
    {
        CellAddress.TryGetColumnRow(id, out var col, out var row);
        return new CellDep(sheet, col, row);
    }

    // --- Cenários controlados ---

    [Test]
    public async Task SimpleChain_EditPropagatesUp()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s);
        s["A3"] = ExpressionParser.Parse("=A2+1", s);

        var affected = ReverseDependencyGraph.Build(wb).GetAllDependents([Addr("Sheet1", "A1")]);

        await Assert
            .That(affected.SetEquals([Addr("Sheet1", "A2"), Addr("Sheet1", "A3")]))
            .IsTrue();
    }

    [Test]
    public async Task BoundedRange_EditInsideAffectsTheAggregate()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        for (var r = 1; r <= 5; r++)
        {
            s[$"A{r}"] = new NumberValue(r);
        }
        s["B1"] = ExpressionParser.Parse("=SUM(A1:A5)", s);
        s["C1"] = ExpressionParser.Parse("=B1*2", s);

        var affected = ReverseDependencyGraph.Build(wb).GetAllDependents([Addr("Sheet1", "A3")]);

        await Assert
            .That(affected.SetEquals([Addr("Sheet1", "B1"), Addr("Sheet1", "C1")]))
            .IsTrue();
    }

    [Test]
    public async Task WholeColumn_ContainsAnyRowEvenUnpopulated()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = new NumberValue(2);
        s["B1"] = ExpressionParser.Parse("=SUM(A:A)", s);

        var graph = ReverseDependencyGraph.Build(wb);

        // Uma linha populada e uma linha AINDA vazia da coluna A ambas afetam SUM(A:A).
        await Assert
            .That(graph.GetAllDependents([Addr("Sheet1", "A2")]).SetEquals([Addr("Sheet1", "B1")]))
            .IsTrue();
        await Assert
            .That(
                graph.GetAllDependents([Addr("Sheet1", "A9999")]).SetEquals([Addr("Sheet1", "B1")])
            )
            .IsTrue();
        // Uma célula FORA da coluna A não afeta.
        await Assert.That(graph.GetAllDependents([Addr("Sheet1", "C1")]).Count).IsEqualTo(0);
    }

    [Test]
    public async Task CrossSheet_EdgeCarriesTheSheet()
    {
        var wb = new Workbook();
        var s1 = wb.Sheets.Add("Sheet1");
        var s2 = wb.Sheets.Add("Sheet2");
        s1["A1"] = new NumberValue(10);
        s2["B1"] = ExpressionParser.Parse("=Sheet1!A1*2", s2);

        var graph = ReverseDependencyGraph.Build(wb);

        await Assert
            .That(graph.GetAllDependents([Addr("Sheet1", "A1")]).SetEquals([Addr("Sheet2", "B1")]))
            .IsTrue();
        // Editar Sheet2!A1 (mesma célula, outra sheet) não afeta nada.
        await Assert.That(graph.GetAllDependents([Addr("Sheet2", "A1")]).Count).IsEqualTo(0);
    }

    [Test]
    public async Task Diamond_ReachesAllDownstream()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["B1"] = ExpressionParser.Parse("=A1", s);
        s["C1"] = ExpressionParser.Parse("=A1", s);
        s["D1"] = ExpressionParser.Parse("=B1+C1", s);

        var affected = ReverseDependencyGraph.Build(wb).GetAllDependents([Addr("Sheet1", "A1")]);

        await Assert
            .That(
                affected.SetEquals([
                    Addr("Sheet1", "B1"),
                    Addr("Sheet1", "C1"),
                    Addr("Sheet1", "D1"),
                ])
            )
            .IsTrue();
    }

    [Test]
    public async Task Volatile_IsAlwaysDirty_NotADependentOfAnEdit()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["V1"] = ExpressionParser.Parse("=NOW()", s);
        s["B1"] = ExpressionParser.Parse("=A1+1", s);

        var graph = ReverseDependencyGraph.Build(wb);

        // NOW() não é dependente de editar A1 (não lê A1), mas está no conjunto sempre-dirty.
        await Assert
            .That(graph.GetAllDependents([Addr("Sheet1", "A1")]).SetEquals([Addr("Sheet1", "B1")]))
            .IsTrue();
        await Assert.That(graph.AlwaysDirty.Contains(Addr("Sheet1", "V1"))).IsTrue();
    }

    // --- Comparação contra o oráculo num workbook randomizado ---

    [Test]
    public async Task MatchesBruteForceOracle_OnRandomWorkbook()
    {
        var rng = new Random(42);
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");

        // 300 células na coluna A: as 10 primeiras literais; o resto é uma fórmula que lê 1-3 células
        // anteriores (às vezes um SUM de um range anterior), criando um grafo denso e profundo.
        const int n = 300;
        for (var r = 1; r <= 10; r++)
        {
            s[$"A{r}"] = new NumberValue(r);
        }
        for (var r = 11; r <= n; r++)
        {
            if (rng.Next(4) == 0)
            {
                var lo = rng.Next(1, r - 1);
                var hi = rng.Next(lo, r - 1);
                s[$"A{r}"] = ExpressionParser.Parse($"=SUM(A{lo}:A{hi})", s);
            }
            else
            {
                var a = rng.Next(1, r);
                var b = rng.Next(1, r);
                s[$"A{r}"] = ExpressionParser.Parse($"=A{a}+A{b}", s);
            }
        }

        var graph = ReverseDependencyGraph.Build(wb);

        // 40 edições aleatórias: o grafo deve casar com o oráculo bit a bit.
        for (var t = 0; t < 40; t++)
        {
            var edited = new[] { Addr("Sheet1", $"A{rng.Next(1, n + 1)}") };
            var fromGraph = graph.GetAllDependents(edited);
            var fromOracle = Oracle(wb, edited);

            await Assert.That(fromGraph.SetEquals(fromOracle)).IsTrue();
        }
    }

    // Oráculo: fecho transitivo dos dependentes re-derivado varrendo TODAS as fórmulas a cada iteração.
    private static HashSet<CellDep> Oracle(Workbook workbook, IEnumerable<CellDep> edited)
    {
        var formulas = new List<(CellDep Addr, DependencyScan Scan)>();
        foreach (var sheet in workbook.Sheets.Values)
        {
            foreach (var (id, expression) in sheet)
            {
                if (
                    expression is ValueExpression
                    || !CellAddress.TryGetColumnRow(id, out var col, out var row)
                )
                {
                    continue;
                }
                formulas.Add(
                    (
                        new CellDep(sheet.Name, col, row),
                        DependencyExtractor.Extract(expression, workbook)
                    )
                );
            }
        }

        var editedSet = new HashSet<CellDep>(edited);
        var affected = new HashSet<CellDep>();

        var changed = true;
        while (changed)
        {
            changed = false;
            foreach (var (addr, scan) in formulas)
            {
                if (affected.Contains(addr) || !ReadsAny(scan, editedSet, affected))
                {
                    continue;
                }
                affected.Add(addr);
                changed = true;
            }
        }

        return affected;
    }

    private static bool ReadsAny(
        DependencyScan scan,
        HashSet<CellDep> edited,
        HashSet<CellDep> affected
    )
    {
        foreach (var source in scan.Cells)
        {
            if (edited.Contains(source) || affected.Contains(source))
            {
                return true;
            }
        }

        foreach (var range in scan.Ranges)
        {
            if (RangeHitsAny(range, edited) || RangeHitsAny(range, affected))
            {
                return true;
            }
        }

        return false;
    }

    private static bool RangeHitsAny(RangeDep range, HashSet<CellDep> cells)
    {
        foreach (var cell in cells)
        {
            if (
                string.Equals(range.Sheet, cell.Sheet, StringComparison.OrdinalIgnoreCase)
                && (range.ColMin is not { } cMin || cell.Column >= cMin)
                && (range.ColMax is not { } cMax || cell.Column <= cMax)
                && (range.RowMin is not { } rMin || cell.Row >= rMin)
                && (range.RowMax is not { } rMax || cell.Row <= rMax)
            )
            {
                return true;
            }
        }

        return false;
    }
}
