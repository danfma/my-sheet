using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Tests.DirtyGraph;

// Fase 7 — SEGUNDO WORKLOAD (corretude num shape estruturalmente diferente do K1). O K1 é quase mono-sheet com
// UMA coluna quente; aqui o workbook é uma MALHA de 4 sheets com dependências cross-sheet pesadas, cadeias
// profundas atravessando sheets, ranges abertos + fechados cross-sheet, uniões e ramos de IF. O portão é o
// mesmo: lotes que misturam edições de VALOR com edições de FÓRMULA (incluindo RE-WIRING cross-sheet e
// conversões literal↔fórmula) devem ser BIT-IDÊNTICOS a InvalidateCache()+ComputeAll(). Se o rebuild lazy
// perdesse uma dep cross-sheet nova, um HIT stale divergiria.
public class CrossSheetFormulaEditEquivalenceTests
{
    private const int Rows = 80;
    private const int MutableCells = 8;

    // Raw (inputs) → Calc (cadeias + ranges cross-sheet) → Agg (open range + união cross-sheet) → Report
    // (sinks + células M MUTÁVEIS que reescrevemos por lote, com deps que atravessam sheets).
    private static Workbook BuildMesh()
    {
        var wb = new Workbook();
        var raw = wb.Sheets.Add("Raw");
        var calc = wb.Sheets.Add("Calc");
        var agg = wb.Sheets.Add("Agg");
        var report = wb.Sheets.Add("Report");

        for (var r = 1; r <= Rows; r++)
        {
            raw[$"A{r}"] = new NumberValue(r);
            raw[$"B{r}"] = new NumberValue(Rows - r + 1);
        }

        for (var r = 1; r <= Rows; r++)
        {
            // cadeia profunda + produto cross-sheet
            calc[$"A{r}"] = ExpressionParser.Parse(
                r == 1 ? "=Raw!A1*Raw!B1" : $"=Raw!A{r}*Raw!B{r}+Calc!A{r - 1}",
                calc
            );
            // range fechado cross-sheet (prefixo acumulado) menos um input
            calc[$"B{r}"] = ExpressionParser.Parse($"=SUM(Raw!A1:A{r})-Raw!B{r}", calc);
            // IF (ambos os ramos coletados) sobre células cross-sheet
            calc[$"C{r}"] = ExpressionParser.Parse(
                $"=IF(Calc!A{r}>Calc!B{r},Raw!A{r},Raw!B{r})",
                calc
            );
        }

        agg["A1"] = ExpressionParser.Parse($"=SUM(Calc!A1:A{Rows})", agg); // range fechado cross-sheet
        agg["A2"] = ExpressionParser.Parse("=SUM(Calc!B:B)", agg); // coluna ABERTA cross-sheet
        agg["A3"] = ExpressionParser.Parse("=SUM((Calc!A1:A10,Calc!C1:C10))", agg); // UNIÃO cross-sheet
        agg["A4"] = ExpressionParser.Parse("=Agg!A1+Agg!A2-Agg!A3", agg); // agregado de agregados

        report["R1"] = ExpressionParser.Parse("=Agg!A4*2", report);
        report["R2"] = ExpressionParser.Parse("=Agg!A1-Agg!A3", report);

        // Células MUTÁVEIS (fórmula inicial simples cross-sheet).
        for (var i = 1; i <= MutableCells; i++)
        {
            report[$"M{i}"] = ExpressionParser.Parse($"=Raw!A{i}", report);
        }

        // Topo que lê o range das M mutáveis (uma edição de fórmula em M propaga até aqui).
        report["T1"] = ExpressionParser.Parse($"=SUM(Report!M1:M{MutableCells})+Agg!A4", report);

        return wb;
    }

    // Fórmula aleatória para Report!M{i}: referencia só Raw (inputs), Calc!A (que só lê Raw+Calc!A, sem Report)
    // e Report!M{j<i} → sempre acíclica, sempre CROSS-SHEET.
    private static string RandomMeshFormula(int i, Random rng)
    {
        var a = rng.Next(1, Rows + 1);
        var b = rng.Next(1, Rows + 1);
        var coef = rng.Next(1, 5);
        var formula = $"=Raw!A{a}+{coef}*Raw!B{b}";
        if (rng.Next(2) == 0)
        {
            formula += $"-Calc!A{rng.Next(1, Rows + 1)}";
        }
        if (i > 1 && rng.Next(2) == 0)
        {
            formula += $"+Report!M{rng.Next(1, i)}"; // um M estritamente menor
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
    public async Task Recalculate_IsBitIdenticalToFullRecompute_OnACrossSheetMesh()
    {
        var wb = BuildMesh();
        var raw = wb["Raw"];
        var report = wb["Report"];
        var engine = wb.CreateRecalculationEngine();

        wb.GetCellValue("Report", "T1"); // aquece

        var rng = new Random(20260709);
        var sawFormulaEdit = false;
        var sawValueOnly = false;

        for (var iteration = 0; iteration < 80; iteration++)
        {
            var edited = new List<CellRef>();
            var versionBefore = report.StructuralVersion;

            // (a) 0..3 edições de VALOR em inputs Raw (colunas A e B).
            var valueEdits = rng.Next(0, 4);
            for (var i = 0; i < valueEdits; i++)
            {
                var col = rng.Next(2) == 0 ? "A" : "B";
                var id = $"{col}{rng.Next(1, Rows + 1)}";
                raw[id] = new NumberValue(rng.Next(-1000, 1000));
                edited.Add(new CellRef("Raw", id));
            }

            // (b) 0..2 edições de FÓRMULA/conversões em Report!M (re-wiring cross-sheet).
            var formulaEdits = rng.Next(0, 3);
            for (var i = 0; i < formulaEdits; i++)
            {
                var target = rng.Next(1, MutableCells + 1);
                var id = $"M{target}";
                switch (rng.Next(3))
                {
                    case 0: // reescreve para outra fórmula cross-sheet (deps diferentes)
                    case 2: // idem (mais peso para edição de fórmula)
                        report[id] = ExpressionParser.Parse(RandomMeshFormula(target, rng), report);
                        break;
                    default: // fórmula → literal (estrutural)
                        report[id] = new NumberValue(rng.Next(-500, 500));
                        break;
                }
                edited.Add(new CellRef("Report", id));
            }

            if (edited.Count == 0)
            {
                continue;
            }

            var structuralChanged = report.StructuralVersion != versionBefore;

            var result = engine.Recalculate(edited);
            await Assert.That(result.StructureRebuilt).IsEqualTo(structuralChanged);

            var afterDirty = SnapshotAll(wb);

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
                            + $"{string.Join(",", edited.ConvertAll(c => $"{c.Sheet}!{c.Id}"))}. "
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

        await Assert.That(sawFormulaEdit).IsTrue();
        await Assert.That(sawValueOnly).IsTrue();
    }
}
