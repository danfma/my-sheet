using BenchmarkDotNet.Running;
using Danfma.MySheet.Benchmark;
using Danfma.MySheet.Benchmark.Spike;
using Danfma.MySheet.Benchmark.Spike.DirtyGraph;
using Danfma.MySheet.Benchmark.Spike.MessagePackFormat;
using Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// TEMP memória Excel (OpenXML vs Aspose): `dotnet run -c Release -- --excel-memory`.
if (args.Contains("--excel-memory"))
{
    ExcelMemoryHarness.Run();
    return;
}

// Sanidade do spike CellValue (Fase 1): `dotnet run -- --check`.
if (args.Contains("--check"))
{
    SpikeSelfCheck.Run();
    return;
}

// Baseline do spike do grafo de dependências (Fase 0): `dotnet run -c Release -- --dirty-graph [N]`.
if (args.Contains("--dirty-graph"))
{
    DirtyGraphHarness.Run(args);
    return;
}

// Footprint do grafo reverso na K1 (Fase 2): `dotnet run -c Release -- --dep-graph`.
if (args.Contains("--dep-graph"))
{
    DirtyGraphHarness.RunGraphFootprint();
    return;
}

// Speedup end-to-end do evict-and-pull na K1 (Fase 4): `dotnet run -c Release -- --dirty-recompute`.
if (args.Contains("--dirty-recompute"))
{
    DirtyGraphHarness.RunDirtyRecompute();
    return;
}

// Wall-clock harness da escala de coluna inteira (Fase 0): `dotnet run -c Release -- --whole-column-scale [--full]`.
if (args.Contains("--whole-column-scale"))
{
    WholeColumnScaleHarness.Run(args);
    return;
}

// Regressão de admissão do range cache (Fase 4): `dotnet run -c Release -- --range-cache-admission`.
if (args.Contains("--range-cache-admission"))
{
    RangeCacheAdmissionHarness.Run();
    return;
}

// Índice estrutural vitalício do 3.0 (Fase 2): `dotnet run -c Release -- --structural-index-lifetime`.
if (args.Contains("--structural-index-lifetime"))
{
    StructuralIndexLifetimeHarness.Run();
    return;
}

// Gate de custo de escrita do 3.0 (Fase 3): `dotnet run -c Release -- --write-cost`.
if (args.Contains("--write-cost"))
{
    WriteCostHarness.Run();
    return;
}

// Comparativo local MySheet × Aspose.Cells (mesma carga K1, tudo em memória): `dotnet run -c Release -- --aspose-compare`.
if (args.Contains("--aspose-compare"))
{
    AsposeCompareHarness.Run();
    return;
}

// Decomposição por fase MySheet × Aspose na carga K1 (in-memory): `dotnet run -c Release -- --k1-endtoend`.
if (args.Contains("--k1-endtoend"))
{
    K1EndToEndHarness.Run();
    return;
}

// Atribuição de custo do compute K1 (probes dirigidos): `dotnet run -c Release -- --k1-compute-attrib`.
if (args.Contains("--k1-compute-attrib"))
{
    K1ComputeProfileHarness.RunAttribution();
    return;
}

// Loop compute-only da carga K1 para profiler externo: `dotnet run -c Release -- --k1-compute-loop [N]`.
if (args.Contains("--k1-compute-loop"))
{
    K1ComputeProfileHarness.RunLoop(args);
    return;
}

// Atribuição de ALOCAÇÃO do compute K1 (probes dirigidos): `dotnet run -c Release -- --k1-compute-alloc`.
if (args.Contains("--k1-compute-alloc"))
{
    K1ComputeProfileHarness.RunAllocAttribution();
    return;
}

// Spike do store denso paginado (Fase 0): `dotnet run -c Release -- --dense-store-spike`.
if (args.Contains("--dense-store-spike"))
{
    DenseStoreSpikeHarness.Run(args);
    return;
}

// Spike v4 (pós-3.3): persistência do índice + revisita da AST numérica.
//   --v4-index-rebuild | --v4-index-persist | --v4-resident | --v4-parse | --v4-hotpath | --v4-all
if (args.Contains("--v4-all"))
{
    V4IndexAstHarness.RunAll(args);
    return;
}
if (args.Contains("--v4-index-rebuild"))
{
    V4IndexAstHarness.IndexRebuild();
    return;
}
if (args.Contains("--v4-index-persist"))
{
    V4IndexAstHarness.IndexPersist();
    return;
}
if (args.Contains("--v4-resident"))
{
    V4IndexAstHarness.Resident();
    return;
}
if (args.Contains("--v4-parse"))
{
    V4IndexAstHarness.Parse();
    return;
}
if (args.Contains("--v4-hotpath"))
{
    V4IndexAstHarness.HotPath();
    return;
}

// Varredura de tamanho de página do store denso (Fase 1, diretriz do dono):
// `dotnet run -c Release -- --dense-store-pagesize`.
if (args.Contains("--dense-store-pagesize"))
{
    DenseStorePageSizeHarness.Run();
    return;
}

// Adaptive first-page promotion efficacy gate (3.3 Phase 2), driving the REAL store:
// `dotnet run -c Release -- --adaptive-first-page`.
if (args.Contains("--adaptive-first-page"))
{
    AdaptiveFirstPageHarness.Run();
    return;
}

// Numeric structural-index memory + open-range parse probe (3.3 Phase 3):
// `dotnet run -c Release -- --structural-index-memory`.
if (args.Contains("--structural-index-memory"))
{
    StructuralIndexMemoryHarness.Run();
    return;
}

// Range read strategies over the dense paged store (ReadOnlySequence / span visitor / IEnumerable), owner
// question 2026-07-04: `dotnet run -c Release -- --range-sequence-probe`.
if (args.Contains("--range-sequence-probe"))
{
    RangeSequenceProbeHarness.Run(args);
    return;
}

// Custo do caminho mini-CSE array (Fase C): `dotnet run -c Release -- --mini-cse-cost`.
if (args.Contains("--mini-cse-cost"))
{
    MiniCseCostHarness.Run();
    return;
}

// Spike MessagePack format — byte sizes: `dotnet run -c Release -- --messagepack-size`.
if (args.Contains("--messagepack-size"))
{
    MessagePackSizeReport.Run();
    return;
}

// Suíte A/B de pesca de alocação por função (plans/function-allocation-fishing.md, Fase 1):
//   `dotnet run -c Release -- --filter *FunctionBenchmarks* --job short`.
//
// Spike CellValue (Fase 2): `dotnet run -c Release -- --filter *Spike*`.
// Sem filtro cai no menu do BenchmarkSwitcher; SheetBenchmarks/FunctionBenchmarks continuam disponíveis.
BenchmarkSwitcher.FromAssembly(typeof(SheetBenchmarks).Assembly).Run(args);
