using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 7 — PORTÃO DE CORRETUDE PARA EDIÇÃO DE FÓRMULA. O portão da Fase 4 só editava VALORES (grafo estável).
// Aqui, cada lote mistura edições de VALOR com edições de FÓRMULA (troca as deps de uma célula) e conversões
// literal↔fórmula — ou seja, a ESTRUTURA muda em runtime. Após o Recalculate (evict-and-pull, que deve
// reconstruir o grafo quando a estrutura mudou), os valores de TODAS as células devem ser BIT-IDÊNTICOS aos de
// um InvalidateCache()+ComputeAll(). Se o rebuild lazy falhasse (grafo stale após editar fórmula), uma dep
// nova seria perdida → HIT stale → divergência. Usa a API PÚBLICA (RecalculationEngine) de ponta a ponta.
public class FormulaEditRecomputeEquivalenceTests
{
    // A1..A20 literais (inputs); B/C cadeias estáveis (célula + range); M1..M5 fórmulas MUTÁVEIS (reescritas a
    // cada lote, sempre referenciando só A-inputs e M menores → acíclico); T1/D* agregados de topo que leem M e
    // C, para uma mudança de fórmula em M propagar até o topo.
    private static (Workbook Wb, Sheet Sheet) BuildWeb()
    {
        var wb = new Workbook();
        var s = wb.Sheets.Add("Sheet1");

        for (var r = 1; r <= 20; r++)
        {
            s[$"A{r}"] = new NumberValue(r);
        }

        for (var r = 1; r <= 20; r++)
        {
            s[$"B{r}"] = ExpressionParser.Parse(r == 1 ? "=A1*2" : $"=A{r}*2+B{r - 1}", s);
            s[$"C{r}"] = ExpressionParser.Parse($"=SUM(A1:A{r})+B{r}", s);
        }

        for (var i = 1; i <= 5; i++)
        {
            s[$"M{i}"] = ExpressionParser.Parse($"=A{i}", s); // fórmula inicial simples
        }

        s["D1"] = ExpressionParser.Parse("=SUM(C1:C20)", s);
        s["D2"] = ExpressionParser.Parse("=SUM(B:B)-D1", s);
        s["T1"] = ExpressionParser.Parse("=SUM(M1:M5)+D1-D2", s); // topo que lê o range de M mutável

        return (wb, s);
    }

    // Uma fórmula aleatória para M{i}: referencia só A1..A20 e (opcionalmente) um M{j} com j < i → nunca cíclica.
    private static string RandomMFormula(int i, Random rng)
    {
        var a = rng.Next(1, 21);
        var b = rng.Next(1, 21);
        var coef = rng.Next(1, 5);
        var formula = $"=A{a}+{coef}*A{b}";
        if (i > 1 && rng.Next(2) == 0)
        {
            formula += $"-M{rng.Next(1, i)}"; // um M estritamente menor (acíclico)
        }
        return formula;
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
    public async Task Recalculate_IsBitIdenticalToFullRecompute_AcrossFormulaEdits()
    {
        var (wb, sheet) = BuildWeb();
        var engine = wb.CreateRecalculationEngine();

        wb.GetCellValue("Sheet1", "T1"); // aquece

        var rng = new Random(20260708);
        var sawFormulaEdit = false;
        var sawValueOnly = false;

        for (var iteration = 0; iteration < 80; iteration++)
        {
            var edited = new List<CellRef>();
            var versionBefore = sheet.StructuralVersion;

            // (a) 0..2 edições de VALOR em inputs A.
            var valueEdits = rng.Next(0, 3);
            for (var i = 0; i < valueEdits; i++)
            {
                var id = $"A{rng.Next(1, 21)}";
                sheet[id] = new NumberValue(rng.Next(-1000, 1000));
                edited.Add(new CellRef("Sheet1", id));
            }

            // (b) 0..2 edições de FÓRMULA/conversões em M — muda a ESTRUTURA (deps).
            var formulaEdits = rng.Next(0, 3);
            for (var i = 0; i < formulaEdits; i++)
            {
                var target = rng.Next(1, 6); // M1..M5
                var id = $"M{target}";
                var current = sheet[id];

                switch (rng.Next(3))
                {
                    case 0: // reescreve para outra fórmula (deps diferentes)
                        sheet[id] = ExpressionParser.Parse(RandomMFormula(target, rng), sheet);
                        break;
                    case 1: // fórmula → literal (estrutural: descarta as arestas de saída)
                        sheet[id] = new NumberValue(rng.Next(-500, 500));
                        break;
                    default: // literal → fórmula, ou fórmula → outra fórmula
                        sheet[id] = ExpressionParser.Parse(RandomMFormula(target, rng), sheet);
                        break;
                }

                _ = current;
                edited.Add(new CellRef("Sheet1", id));
            }

            if (edited.Count == 0)
            {
                continue; // lote vazio (raro): nada a comparar
            }

            var structuralChanged = sheet.StructuralVersion != versionBefore;

            // Evict-and-pull via a MESMA instância do engine (persiste entre lotes); deve reconstruir o grafo
            // internamente sse a estrutura mudou neste lote.
            var result = engine.Recalculate(edited);

            // O rebuild aconteceu EXATAMENTE quando a estrutura mudou neste lote.
            await Assert.That(result.StructureRebuilt).IsEqualTo(structuralChanged);

            var afterDirty = SnapshotAll(wb);

            // Referência: full recompute do MESMO estado.
            wb.InvalidateCache();
            wb.ComputeAll();
            var afterFull = SnapshotAll(wb);

            await Assert.That(afterDirty.Count).IsEqualTo(afterFull.Count);
            foreach (var (key, dirtyValue) in afterDirty)
            {
                var fullValue = afterFull[key];
                if (!Equals(dirtyValue, fullValue))
                {
                    Assert.Fail(
                        $"Iteração {iteration}: {key} divergiu (evict-and-pull={dirtyValue ?? "blank"}, "
                            + $"full={fullValue ?? "blank"}). Editadas: "
                            + $"{string.Join(",", edited.ConvertAll(c => c.Id))}. "
                            + $"Estrutura mudou: {structuralChanged}; modo: {result.Mode}."
                    );
                }
            }

            if (structuralChanged)
            {
                sawFormulaEdit = true;
            }
            else
            {
                sawValueOnly = true;
            }
        }

        // Sanidade: o fuzz exercitou AMBOS os caminhos (rebuild e não-rebuild).
        await Assert.That(sawFormulaEdit).IsTrue();
        await Assert.That(sawValueOnly).IsTrue();
    }
}
