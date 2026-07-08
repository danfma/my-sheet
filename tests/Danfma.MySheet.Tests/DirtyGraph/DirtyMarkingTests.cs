using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 3: marcação dirty + invalidação seletiva (evict-and-pull), sem tocar InvalidateCache.
public class DirtyMarkingTests
{
    private static CellDep Addr(Workbook _, string sheet, string id)
    {
        CellAddress.TryGetColumnRow(id, out var col, out var row);
        return new CellDep(sheet, col, row);
    }

    [Test]
    public async Task CalculateDirty_IsEditedPlusTransitiveDependents()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s);
        s["A3"] = ExpressionParser.Parse("=A2+1", s);

        var engine = DirtyEngine.Build(wb);
        var dirty = engine.CalculateDirty([Addr(wb, "Sheet1", "A1")]);

        // A1 (editada) + A2 + A3 (dependentes transitivos).
        await Assert
            .That(
                dirty!.SetEquals([
                    Addr(wb, "Sheet1", "A1"),
                    Addr(wb, "Sheet1", "A2"),
                    Addr(wb, "Sheet1", "A3"),
                ])
            )
            .IsTrue();
    }

    [Test]
    public async Task EvictAndPull_PropagatesEdit_WithoutInvalidateCache()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s);
        s["A3"] = ExpressionParser.Parse("=A2+1", s);

        var engine = DirtyEngine.Build(wb);

        // Aquece o cache: A3 = 3.
        await Assert.That(wb.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(3.0);

        // Edita A1 (SetCell NÃO invalida o cache — A3 continuaria stale=3 sem o evict).
        s["A1"] = new NumberValue(10);
        engine.CalculateDirty([Addr(wb, "Sheet1", "A1")]);

        // Sem NENHUM InvalidateCache: A3 recomputa via pull para 12.
        await Assert.That(wb.GetCellValue("Sheet1", "A3").ToDouble()).IsEqualTo(12.0);
    }

    [Test]
    public async Task Evict_IsSelective_UnaffectedCellsStayCached()
    {
        // Uma célula fora do cone dirty NÃO entra no conjunto dirty → não é evictada → permanece um HIT
        // cacheado. Prova via a composição do conjunto (B1 lê Z1, não A1, então editar A1 não o toca).
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s); // depende de A1
        s["Z1"] = new NumberValue(100);
        s["B1"] = ExpressionParser.Parse("=Z1+1", s); // depende de Z1, NÃO de A1

        var engine = DirtyEngine.Build(wb);
        var dirty = engine.CalculateDirty([Addr(wb, "Sheet1", "A1")]);

        await Assert.That(dirty!.Contains(Addr(wb, "Sheet1", "A2"))).IsTrue(); // afetada
        await Assert.That(dirty!.Contains(Addr(wb, "Sheet1", "B1"))).IsFalse(); // não evictada → HIT
        await Assert.That(dirty!.Contains(Addr(wb, "Sheet1", "Z1"))).IsFalse();
    }

    [Test]
    public async Task GetAffectedOutputs_AreTheSinks()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["A2"] = ExpressionParser.Parse("=A1+1", s);
        s["A3"] = ExpressionParser.Parse("=A2+1", s); // sink
        s["B1"] = ExpressionParser.Parse("=A1*2", s); // sink (nada dirty depende dele)

        var engine = DirtyEngine.Build(wb);
        var dirty = engine.CalculateDirty([Addr(wb, "Sheet1", "A1")]);
        var outputs = engine.GetAffectedOutputs(dirty!);

        // Sinks = células dirty sem dependente dirty: A3 e B1 (A1 e A2 têm dependentes dirty).
        await Assert
            .That(
                new HashSet<CellDep>(outputs).SetEquals([
                    Addr(wb, "Sheet1", "A3"),
                    Addr(wb, "Sheet1", "B1"),
                ])
            )
            .IsTrue();
    }

    [Test]
    public async Task AlwaysDirty_IsAlwaysInTheDirtySet()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        s["A1"] = new NumberValue(1);
        s["B1"] = ExpressionParser.Parse("=A1+1", s);
        s["V1"] = ExpressionParser.Parse("=NOW()", s); // sempre-dirty

        var engine = DirtyEngine.Build(wb);
        var dirty = engine.CalculateDirty([Addr(wb, "Sheet1", "A1")]);

        // V1 não depende de A1, mas é sempre-dirty → sempre no conjunto.
        await Assert.That(dirty!.Contains(Addr(wb, "Sheet1", "V1"))).IsTrue();
    }
}
