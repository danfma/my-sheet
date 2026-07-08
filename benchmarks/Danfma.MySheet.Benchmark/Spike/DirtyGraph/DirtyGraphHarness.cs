using System.Diagnostics;
using System.Runtime;
using Danfma.MySheet;
using Danfma.MySheet.DirtyGraph;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike.DirtyGraph;

/// <summary>
/// Baseline do spike do grafo de dependências (Fase 0). Carrega o fixture K1 anonimizado (samples/k1.myxl,
/// 566k células) e mede o custo ATUAL de uma edição pontual no modelo tudo-ou-nada: sem grafo reverso, a
/// única forma de obter resultados frescos após editar uma célula é <c>InvalidateCache()</c> + recomputar
/// tudo. Este é o número que o CalculateDirty (Fases 3–4) precisa bater.
///
/// Uso: <c>dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --dirty-graph [N]</c>
/// (N = nº de edições pontuais medidas; default 8).
/// </summary>
internal static class DirtyGraphHarness
{
    public static void Run(string[] args)
    {
        var edits = 8;
        var nArg = Array.FindIndex(args, a => a == "--dirty-graph");
        if (nArg >= 0 && nArg + 1 < args.Length && int.TryParse(args[nArg + 1], out var parsed))
        {
            edits = parsed;
        }

        var fixturePath = FindFixture();
        if (fixturePath is null)
        {
            Console.Error.WriteLine(
                "Fixture samples/k1.myxl não encontrado. Gere com tools/K1FixtureBuilder."
            );
            return;
        }

        Console.WriteLine("=== Dirty-graph spike — Fase 0 (baseline) ===");
        Console.WriteLine(
            $"Fixture: {fixturePath} ({new FileInfo(fixturePath).Length / 1_000_000.0:F1} MB)"
        );

        // --- Load ---
        var sw = Stopwatch.StartNew();
        var workbook = Workbook.Load(fixturePath);
        sw.Stop();
        var loadMs = sw.Elapsed.TotalMilliseconds;

        // Lista plana de células populadas (para escolher alvos de edição) + contagem.
        var cells = new List<(string Sheet, string Id)>();
        foreach (var sheet in workbook.Sheets.Values)
        {
            foreach (var id in sheet.Keys)
            {
                cells.Add((sheet.Name, id));
            }
        }
        Console.WriteLine(
            $"Sheets: {workbook.Sheets.Count} | células populadas: {cells.Count} | load: {loadMs:F0} ms"
        );
        Console.WriteLine();

        // --- Cold ComputeAll (primeira calculação completa) ---
        var (coldMs, coldAlloc) = Measure(() => workbook.ComputeAll());
        Console.WriteLine(
            $"ComputeAll frio (1ª calc completa): {coldMs:F0} ms, {coldAlloc / 1_000_000.0:F1} MB alloc"
        );
        Console.WriteLine();

        // --- Baseline: N edições pontuais, cada uma = editar + InvalidateCache + ComputeAll ---
        // Semente fixa p/ reprodutibilidade; alvos aleatórios dentre as células populadas.
        var rng = new Random(20260708);
        var times = new double[edits];
        var allocs = new long[edits];

        Console.WriteLine(
            $"Baseline de edição pontual (N={edits}): editar 1 célula → InvalidateCache → ComputeAll"
        );
        for (var k = 0; k < edits; k++)
        {
            var (sheetName, id) = cells[rng.Next(cells.Count)];
            var sheet = workbook[sheetName];
            var original = sheet[id]; // captura a expressão original p/ restaurar depois

            var (ms, alloc) = Measure(() =>
            {
                sheet[id] = new NumberValue(rng.Next(1, 1_000_000));
                workbook.InvalidateCache();
                workbook.ComputeAll();
            });

            times[k] = ms;
            allocs[k] = alloc;

            // Restaura a célula original (não medido) p/ manter o fixture estável entre iterações.
            sheet[id] = original;
            workbook.InvalidateCache();

            Console.WriteLine(
                $"  edição {k + 1, 2}/{edits}: {sheetName}!{id, -8} → {ms, 7:F0} ms, {alloc / 1_000_000.0, 6:F1} MB"
            );
        }

        Array.Sort(times);
        var mean = times.Average();
        var median =
            edits % 2 == 1 ? times[edits / 2] : (times[edits / 2 - 1] + times[edits / 2]) / 2;

        Console.WriteLine();
        Console.WriteLine(
            "=== BASELINE (custo de UMA edição pontual, modelo atual tudo-ou-nada) ==="
        );
        Console.WriteLine(
            $"  tempo:  média {mean:F0} ms | mediana {median:F0} ms | min {times[0]:F0} | max {times[^1]:F0}"
        );
        Console.WriteLine($"  alloc:  média {allocs.Average() / 1_000_000.0:F1} MB por edição");
        Console.WriteLine();
        Console.WriteLine(
            "Este é o número que CalculateDirty (Fases 3–4) precisa bater para uma edição pontual."
        );
    }

    // Fase 2: constrói o grafo reverso a partir do fixture K1 e reporta o FOOTPRINT real + a distribuição
    // do tamanho dos conjuntos afetados por uma edição pontual (a tese do speedup).
    public static void RunGraphFootprint()
    {
        var fixturePath = FindFixture();
        if (fixturePath is null)
        {
            Console.Error.WriteLine("Fixture samples/k1.myxl não encontrado.");
            return;
        }

        Console.WriteLine("=== Dirty-graph spike — Fase 2 (footprint do grafo reverso) ===");
        var workbook = Workbook.Load(fixturePath);

        // Memória gerenciada RETIDA pelo grafo: mede depois de GC full, antes e depois do build (o workbook
        // já está carregado antes da 1ª medição, então o delta é ~o grafo).
        var memBefore = SettledMemory();

        var sw = Stopwatch.StartNew();
        var graph = ReverseDependencyGraph.Build(workbook);
        sw.Stop();

        var memAfter = SettledMemory();
        GC.KeepAlive(workbook);

        var d = graph.Diagnostics();
        Console.WriteLine($"Build do grafo:          {sw.Elapsed.TotalMilliseconds:F0} ms");
        Console.WriteLine($"Células-fonte distintas: {d.DistinctSourceCells:N0}");
        Console.WriteLine($"Arestas de célula:       {d.CellEdges:N0}");
        Console.WriteLine(
            $"Deps de RANGE (lógicas): {d.RangeDeps:N0}  (uma por fórmula-range — NÃO por célula coberta)"
        );
        Console.WriteLine(
            $"  destas, coluna-inteira: {d.WholeColumnDeps:N0}  (linha aberta — casa qualquer linha da coluna)"
        );
        Console.WriteLine(
            $"Entradas de bucket:      {d.RangeBucketEntries:N0}  (buckets de coluna: {d.ColumnBuckets:N0}, MAIOR bucket: {d.MaxColumnBucket:N0})"
        );
        Console.WriteLine($"Fórmulas sempre-dirty:   {d.AlwaysDirtyFormulas:N0}");
        Console.WriteLine($"Footprint estimado:      {d.EstimatedBytes / 1_000_000.0:F1} MB");
        Console.WriteLine($"Footprint REAL (heap):   {(memAfter - memBefore) / 1_000_000.0:F1} MB");
        Console.WriteLine();

        // Distribuição do tamanho do conjunto afetado por 1 edição pontual + custo da consulta.
        var cells = new List<CellDep>();
        foreach (var sheet in workbook.Sheets.Values)
        {
            foreach (var id in sheet.Keys)
            {
                if (CellAddressTryGet(id, out var col, out var row))
                {
                    cells.Add(new CellDep(sheet.Name, col, row));
                }
            }
        }

        var rng = new Random(20260708);
        var into = new List<CellDep>();

        // Fan-in DIRETO (CollectDirectDependents, custo O(bucket), limitado): quantas fórmulas leem a célula
        // editada diretamente. Isto caracteriza se o fan-in é grande ou pequeno SEM a BFS transitiva (que é
        // quadrática quando o fan-in concentra numa coluna quente — ver o MAIOR bucket acima).
        Console.WriteLine(
            "Fan-in DIRETO por edição pontual (30 edições aleatórias, custo por-query):"
        );
        var directSizes = new int[30];
        var directMs = new double[30];
        for (var i = 0; i < 30; i++)
        {
            var edited = cells[rng.Next(cells.Count)];
            into.Clear();
            var sw2 = Stopwatch.StartNew();
            graph.CollectDirectDependents(edited, into);
            sw2.Stop();
            directSizes[i] = into.Count;
            directMs[i] = sw2.Elapsed.TotalMilliseconds;
        }
        Array.Sort(directSizes);
        Console.WriteLine(
            $"  dependentes diretos: mediana {directSizes[15]:N0} | média {directSizes.Average():F0} | p95 {directSizes[28]:N0} | max {directSizes[^1]:N0}"
        );
        Console.WriteLine(
            $"  custo médio da consulta direta: {directMs.Average():F3} ms/edição (max {directMs.Max():F3} ms)"
        );
        Console.WriteLine();

        // Impacto TRANSITIVO via o sweep fixpoint (GetAllDependents eficiente): total de células afetadas +
        // custo. Compara com as 566k totais — é o teto do que o CalculateDirty recomputaria.
        const int tSamples = 50;
        var totalSizes = new int[tSamples];
        var totalMs = new double[tSamples];
        for (var i = 0; i < tSamples; i++)
        {
            var edited = cells[rng.Next(cells.Count)];
            var sw3 = Stopwatch.StartNew();
            totalSizes[i] = graph.GetAllDependents([edited]).Count;
            sw3.Stop();
            totalMs[i] = sw3.Elapsed.TotalMilliseconds;
        }
        Array.Sort(totalSizes);
        Console.WriteLine(
            $"Impacto TRANSITIVO por edição de célula QUALQUER (N={tSamples}, sweep fixpoint):"
        );
        Console.WriteLine(
            $"  células afetadas: mediana {totalSizes[tSamples / 2]:N0} | média {totalSizes.Average():F0} | p90 {totalSizes[(int)(tSamples * 0.9)]:N0} | max {totalSizes[^1]:N0}  (de 566k)"
        );
        Console.WriteLine(
            $"  custo médio de GetAllDependents: {totalMs.Average():F1} ms/edição (max {totalMs.Max():F1} ms)"
        );
        Console.WriteLine();

        // Caso REALISTA: editar um INPUT (Input!B{r}, r=1..2577) — o cenário "dado inputs → outputs". Pode
        // cascatear bem menos que uma fórmula interna aleatória.
        if (workbook.TryGetSheet("Input", out var input))
        {
            var inputRows = Math.Max(1, input.Count / 3); // ~2577
            var inSizes = new List<int>();
            double inMsTotal = 0;
            var inN = Math.Min(50, inputRows);
            for (var i = 0; i < inN; i++)
            {
                var edited = new CellDep("Input", 2, rng.Next(1, inputRows + 1)); // coluna B
                var sw4 = Stopwatch.StartNew();
                inSizes.Add(graph.GetAllDependents([edited]).Count);
                sw4.Stop();
                inMsTotal += sw4.Elapsed.TotalMilliseconds;
            }
            inSizes.Sort();
            Console.WriteLine(
                $"Impacto TRANSITIVO de editar um INPUT (Input!B, N={inN}) — o caso real:"
            );
            Console.WriteLine(
                $"  células afetadas: mediana {inSizes[inSizes.Count / 2]:N0} | média {inSizes.Average():F0} | p90 {inSizes[(int)(inSizes.Count * 0.9)]:N0} | max {inSizes[^1]:N0}  (de 566k)"
            );
            Console.WriteLine($"  custo médio de GetAllDependents: {inMsTotal / inN:F1} ms/edição");
            Console.WriteLine();
        }
        Console.WriteLine(
            "Leitura: mediana pequena ⇒ o edit TÍPICO recomputaria pouquíssimo (grande ganho do dirty);"
        );
        Console.WriteLine(
            "a cauda (coluna quente, ~490k) é fan-in REAL — ali o dirty não ganha, mas o fixpoint enumera"
        );
        Console.WriteLine(
            "em tempo limitado (não-quadrático). Confirma evict-and-pull como a estratégia da Fase 4."
        );
    }

    // Fase 4: mede o SPEEDUP end-to-end do evict-and-pull (CalculateDirty + pull) vs o baseline
    // (InvalidateCache + ComputeAll), editando INPUTS reais (Input!B) — o caso "dado inputs → outputs".
    public static void RunDirtyRecompute()
    {
        var fixturePath = FindFixture();
        if (fixturePath is null)
        {
            Console.Error.WriteLine("Fixture samples/k1.myxl não encontrado.");
            return;
        }

        Console.WriteLine("=== Dirty-graph spike — Fase 4 (evict-and-pull, speedup real) ===");
        var workbook = Workbook.Load(fixturePath);
        workbook.ComputeAll(); // aquece

        var swBuild = Stopwatch.StartNew();
        var engine = DirtyEngine.Build(workbook);
        swBuild.Stop();
        Console.WriteLine(
            $"Build do motor (grafo), amortizado: {swBuild.Elapsed.TotalMilliseconds:F0} ms (1×)"
        );

        if (!workbook.TryGetSheet("Input", out var input))
        {
            Console.Error.WriteLine("Sheet Input ausente.");
            return;
        }
        var inputRows = Math.Max(1, input.Count / 3);
        var rng = new Random(20260708);

        // Baseline (input edit → InvalidateCache → ComputeAll), poucas amostras p/ referência.
        double baselineMs = 0;
        const int baseN = 3;
        for (var i = 0; i < baseN; i++)
        {
            var r = rng.Next(1, inputRows + 1);
            var original = input[$"B{r}"];
            var (ms, _) = Measure(() =>
            {
                input[$"B{r}"] = new NumberValue(rng.Next(1, 1000));
                workbook.InvalidateCache();
                workbook.ComputeAll();
            });
            baselineMs += ms;
            input[$"B{r}"] = original;
            workbook.InvalidateCache();
            workbook.ComputeAll();
        }
        baselineMs /= baseN;
        Console.WriteLine(
            $"Baseline por edição de input (InvalidateCache+ComputeAll): {baselineMs:F0} ms"
        );
        Console.WriteLine();

        // Dirty (evict-and-pull) por edição de input.
        const int n = 40;
        var smallMs = new List<double>();
        var smallCone = new List<int>();
        var fullFallback = 0;

        for (var k = 0; k < n; k++)
        {
            var r = rng.Next(1, inputRows + 1);
            var addr = new CellDep("Input", 2, r);
            var original = input[$"B{r}"];

            input[$"B{r}"] = new NumberValue(rng.Next(1, 1000));
            var sw = Stopwatch.StartNew();
            var dirty = engine.CalculateDirty([addr]);
            if (dirty is null)
            {
                workbook.ComputeAll(); // cone grande → full-recompute (fallback)
                sw.Stop();
                fullFallback++;
            }
            else
            {
                engine.PullAll(dirty);
                sw.Stop();
                smallMs.Add(sw.Elapsed.TotalMilliseconds);
                smallCone.Add(dirty.Count);
            }

            // restaura (não medido)
            input[$"B{r}"] = original;
            var reset = engine.CalculateDirty([addr]);
            if (reset is null)
            {
                workbook.ComputeAll();
            }
            else
            {
                engine.PullAll(reset);
            }
        }

        Console.WriteLine($"Evict-and-pull por edição de input (N={n}):");
        Console.WriteLine($"  cone PEQUENO (evict-and-pull): {smallMs.Count}/{n} edições");
        Console.WriteLine(
            $"  cone GRANDE (fallback full-recompute): {fullFallback}/{n} edições (cauda da coluna quente)"
        );
        if (smallMs.Count > 0)
        {
            smallMs.Sort();
            smallCone.Sort();
            var medMs = smallMs[smallMs.Count / 2];
            Console.WriteLine(
                $"  cone pequeno — células recomputadas: mediana {smallCone[smallCone.Count / 2]:N0} | max {smallCone[^1]:N0}"
            );
            Console.WriteLine(
                $"  cone pequeno — tempo: mediana {medMs:F2} ms | média {smallMs.Average():F2} ms | max {smallMs.Max():F2} ms"
            );
            Console.WriteLine(
                $"  SPEEDUP (mediana) vs baseline {baselineMs:F0} ms: ~{baselineMs / Math.Max(medMs, 0.001):F0}×"
            );
        }
        Console.WriteLine();

        // API DE ANÁLISE (EstimateImpact): prevê full vs parcial SEM recomputar nem evictar — o host chama
        // para planejar. É BARATO mesmo no caso full (desiste na fonte quente sem varrer o bucket gigante).
        var estMs = new double[100];
        var partial = 0;
        var full = 0;
        for (var i = 0; i < 100; i++)
        {
            var addr = new CellDep("Input", 2, rng.Next(1, inputRows + 1));
            var swEst = Stopwatch.StartNew();
            var estimate = engine.EstimateImpact([addr]);
            swEst.Stop();
            estMs[i] = swEst.Elapsed.TotalMilliseconds;
            if (estimate.RecommendFull)
            {
                full++;
            }
            else
            {
                partial++;
            }
        }
        Array.Sort(estMs);
        Console.WriteLine("API EstimateImpact — prever full/parcial SEM computar (N=100 inputs):");
        Console.WriteLine($"  veredito: {partial}/100 parcial, {full}/100 full");
        Console.WriteLine(
            $"  custo da análise: mediana {estMs[50]:F3} ms | média {estMs.Average():F3} ms | max {estMs[^1]:F3} ms"
        );
        Console.WriteLine(
            "  (barato até no caso full — desiste ao alcançar a fonte quente, sem varrer o bucket de 490k)"
        );
        Console.WriteLine();
        Console.WriteLine(
            "Corretude do evict-and-pull provada no teste DirtyRecomputeEquivalenceTests (bit-idêntico ao full)."
        );
        Console.WriteLine();

        // === Fase 7: custo de uma edição de FÓRMULA (lazy-rebuild) via a API PÚBLICA ====================
        // Uma edição de VALOR fica no caminho rápido (sem rebuild); uma edição de FÓRMULA muda a estrutura, e o
        // engine reconstrói o grafo transparentemente. O objetivo é medir esse custo no fixture real e mostrar
        // que ele é pago SÓ em edição de fórmula (raro), amortizado sobre as muitas edições de valor.
        Console.WriteLine(
            "=== Fase 7 — edição de FÓRMULA (lazy-rebuild) via RecalculationEngine público ==="
        );
        var publicEngine = workbook.CreateRecalculationEngine();

        // (a) edição de VALOR: caminho rápido, StructureRebuilt=false.
        {
            var r = rng.Next(1, inputRows + 1);
            var original = input[$"B{r}"];
            input[$"B{r}"] = new NumberValue(rng.Next(1, 1000));
            var swValue = Stopwatch.StartNew();
            var res = publicEngine.Recalculate([new CellRef("Input", $"B{r}")]);
            swValue.Stop();
            Console.WriteLine(
                $"  edição de VALOR (Input!B{r}): {swValue.Elapsed.TotalMilliseconds:F2} ms | rebuild={res.StructureRebuilt} | modo={res.Mode}"
            );
            input[$"B{r}"] = original;
            publicEngine.Recalculate([new CellRef("Input", $"B{r}")]);
        }

        // (b) edição de FÓRMULA: muda a estrutura → o engine reconstrói o grafo (StructureRebuilt=true).
        var target = FindFormulaCell(workbook);
        if (target is { } t)
        {
            var sheet = workbook[t.Sheet];
            var original = sheet[t.Id];
            sheet[t.Id] = ExpressionParser.Parse("=Input!B1+1", sheet); // deps novas
            var swFormula = Stopwatch.StartNew();
            var res = publicEngine.Recalculate([new CellRef(t.Sheet, t.Id)]);
            swFormula.Stop();
            Console.WriteLine(
                $"  edição de FÓRMULA ({t.Sheet}!{t.Id}): {swFormula.Elapsed.TotalMilliseconds:F0} ms | rebuild={res.StructureRebuilt} | modo={res.Mode}"
            );
            Console.WriteLine(
                $"    (o custo ≈ reconstruir o grafo do zero, ~{swBuild.Elapsed.TotalMilliseconds:F0} ms; pago SÓ em edição de fórmula, amortizado sobre as edições de valor)"
            );
            sheet[t.Id] = original;
            publicEngine.Recalculate([new CellRef(t.Sheet, t.Id)]);
        }
    }

    // Primeira célula de fórmula (não-literal) fora da sheet Input — um alvo seguro para o probe de edição de
    // fórmula.
    private static (string Sheet, string Id)? FindFormulaCell(Workbook workbook)
    {
        foreach (var sheet in workbook.Sheets.Values)
        {
            if (string.Equals(sheet.Name, "Input", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            foreach (var id in sheet.Keys)
            {
                if (sheet[id] is not ValueExpression)
                {
                    return (sheet.Name, id);
                }
            }
        }
        return null;
    }

    // Memória gerenciada viva após GC full + compactação de LOH (o delta de load deixa buffers grandes na
    // LOH que, sem compactar, tornam GetTotalMemory ruidoso — daí o número negativo antes).
    private static long SettledMemory()
    {
        GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    private static bool CellAddressTryGet(string id, out int col, out int row)
    {
        col = 0;
        row = 0;
        var i = 0;
        while (i < id.Length && char.IsLetter(id[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(id[i]) - 'A' + 1);
            i++;
        }
        if (i == 0 || i == id.Length)
        {
            return false;
        }
        var r = 0;
        for (; i < id.Length; i++)
        {
            if (id[i] is < '0' or > '9')
            {
                return false;
            }
            r = r * 10 + (id[i] - '0');
        }
        row = r;
        return true;
    }

    // Mede tempo (ms) e alocação (bytes) de uma ação, coletando lixo antes p/ isolar a medição de alocação.
    private static (double Ms, long Alloc) Measure(Action action)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        var sw = Stopwatch.StartNew();
        action();
        sw.Stop();
        var after = GC.GetTotalAllocatedBytes(precise: true);

        return (sw.Elapsed.TotalMilliseconds, after - before);
    }

    private static string? FindFixture()
    {
        var dir = AppContext.BaseDirectory;
        while (dir is not null && !File.Exists(Path.Combine(dir, "Danfma.MySheet.slnx")))
        {
            dir = Directory.GetParent(dir)?.FullName;
        }

        if (dir is null)
        {
            return null;
        }

        var path = Path.Combine(dir, "samples", "k1.myxl");
        return File.Exists(path) ? path : null;
    }
}
