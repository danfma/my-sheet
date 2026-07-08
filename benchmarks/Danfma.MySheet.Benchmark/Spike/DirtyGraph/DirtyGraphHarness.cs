using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;

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
