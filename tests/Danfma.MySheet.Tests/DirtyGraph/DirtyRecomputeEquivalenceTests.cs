using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 4 — PORTÃO DE CORRETUDE. Sobre um workbook não-trivial (células + ranges + cadeias profundas), K lotes
// de edições aleatórias de VALOR: os valores de TODAS as células após o evict-and-pull (CalculateDirty) devem
// ser BIT-IDÊNTICOS aos após um InvalidateCache()+ComputeAll() do mesmo estado. Se o evict deixou algum valor
// stale (dep perdida), um HIT desatualizado diverge do full-recompute e o teste falha.
public class DirtyRecomputeEquivalenceTests
{
    // Workbook denso: A1..A20 literais (inputs); B/C/D fórmulas com cadeias, deps de célula e de RANGE.
    private static (Workbook Wb, Sheet Sheet, List<string> Inputs) BuildWeb()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");
        var inputs = new List<string>();

        for (var r = 1; r <= 20; r++)
        {
            s[$"A{r}"] = new NumberValue(r);
            inputs.Add($"A{r}");
        }

        for (var r = 1; r <= 20; r++)
        {
            // cadeia + dep de célula
            s[$"B{r}"] = ExpressionParser.Parse(r == 1 ? "=A1*2" : $"=A{r}*2+B{r - 1}", s);
            // dep de RANGE (prefixo acumulado) + célula
            s[$"C{r}"] = ExpressionParser.Parse($"=SUM(A1:A{r})+B{r}", s);
        }

        // agregados de topo (coluna inteira + range fechado)
        s["D1"] = ExpressionParser.Parse("=SUM(C1:C20)", s);
        s["D2"] = ExpressionParser.Parse("=SUM(B:B)-D1", s);
        s["D3"] = ExpressionParser.Parse("=IF(D1>D2,D1,D2)+SUM(A:A)", s);

        return (wb, s, inputs);
    }

    private static CellDep Addr(string id)
    {
        CellAddress.TryGetColumnRow(id, out var col, out var row);
        return new CellDep("Sheet1", col, row);
    }

    private static Dictionary<string, object?> SnapshotAll(Workbook wb)
    {
        var snapshot = new Dictionary<string, object?>();
        foreach (var sheet in wb.Sheets.Values)
        {
            foreach (var id in sheet.Keys)
            {
                snapshot[$"{sheet.Name}!{id}"] = wb.GetCellValue(sheet.Name, id).AsObject();
            }
        }
        return snapshot;
    }

    [Test]
    public async Task CalculateDirty_IsBitIdenticalToFullRecompute_OverRandomEdits()
    {
        var (wb, sheet, inputs) = BuildWeb();
        var engine = DirtyEngine.Build(wb); // grafo do estado inicial; edições são só de VALOR → estável

        wb.GetCellValue("Sheet1", "D3"); // aquece

        var rng = new Random(20260708);

        for (var iteration = 0; iteration < 60; iteration++)
        {
            // 1..3 inputs editados por lote.
            var editedIds = new HashSet<string>();
            var count = rng.Next(1, 4);
            for (var i = 0; i < count; i++)
            {
                editedIds.Add(inputs[rng.Next(inputs.Count)]);
            }

            var editedAddrs = new List<CellDep>();
            foreach (var id in editedIds)
            {
                sheet[id] = new NumberValue(rng.Next(-1000, 1000));
                editedAddrs.Add(Addr(id));
            }

            // Evict-and-pull: evicta o cone dirty; o SnapshotAll a seguir puxa (recomputa) o que foi evictado.
            var dirty = engine.CalculateDirty(editedAddrs);
            var afterDirty = SnapshotAll(wb);

            // Referência: full recompute do MESMO estado.
            wb.InvalidateCache();
            wb.ComputeAll();
            var afterFull = SnapshotAll(wb);

            // Bit-idêntico, célula a célula.
            await Assert.That(afterDirty.Count).IsEqualTo(afterFull.Count);
            foreach (var (key, dirtyValue) in afterDirty)
            {
                var fullValue = afterFull[key];
                if (!Equals(dirtyValue, fullValue))
                {
                    Assert.Fail(
                        $"Iteração {iteration}: {key} divergiu (evict-and-pull={dirtyValue ?? "blank"}, "
                            + $"full={fullValue ?? "blank"}). Editadas: {string.Join(",", editedIds)}. "
                            + $"Cone dirty: {dirty?.Count ?? -1} células (null = fallback full-recompute)."
                    );
                }
            }
        }
    }
}
