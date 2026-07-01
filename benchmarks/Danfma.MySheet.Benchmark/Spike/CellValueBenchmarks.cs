using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;

namespace Danfma.MySheet.Benchmark.Spike;

// Fase 2: compara object? (baseline) vs box-cache (plano-B) vs CellValue nos folds largos (SumFold/Mixed).
// O _Object é Baseline → o BenchmarkDotNet emite a coluna Ratio (delta de throughput do gate). MemoryDiagnoser
// emite Allocated (a coluna que decide "alocação ~0"). ShortRunJob mantém o spike rápido; a alocação é medida
// de forma determinística mesmo em run curto (se o Mean ficar perto do gate de 15%, re-rodar com job completo).
[ShortRunJob]
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class FoldBenchmarks
{
    [Params(1_000, 100_000)]
    public int N;

    private INodeObj _sumObj = null!;
    private INodeObj _sumBoxCache = null!;
    private INodeCv _sumCellValue = null!;
    private INodeObj _mixedObj = null!;
    private INodeObj _mixedBoxCache = null!;
    private INodeCv _mixedCellValue = null!;

    [GlobalSetup]
    public void Setup()
    {
        // Árvores montadas uma vez fora da medição — o benchmark mede só Eval().
        _sumObj = SpikeTrees.SumFoldObj(N);
        _sumBoxCache = SpikeTreesBoxCache.SumFold(N);
        _sumCellValue = SpikeTrees.SumFoldCv(N);
        _mixedObj = SpikeTrees.MixedObj(N);
        _mixedBoxCache = SpikeTreesBoxCache.Mixed(N);
        _mixedCellValue = SpikeTrees.MixedCv(N);
    }

    [Benchmark(Baseline = true), BenchmarkCategory("SumFold")]
    public object? SumFold_Object() => _sumObj.Eval();

    [Benchmark, BenchmarkCategory("SumFold")]
    public object? SumFold_BoxCache() => _sumBoxCache.Eval();

    [Benchmark, BenchmarkCategory("SumFold")]
    public CellValue SumFold_CellValue() => _sumCellValue.Eval();

    [Benchmark(Baseline = true), BenchmarkCategory("Mixed")]
    public object? Mixed_Object() => _mixedObj.Eval();

    [Benchmark, BenchmarkCategory("Mixed")]
    public object? Mixed_BoxCache() => _mixedBoxCache.Eval();

    [Benchmark, BenchmarkCategory("Mixed")]
    public CellValue Mixed_CellValue() => _mixedCellValue.Eval();
}

// Cadeia cumulativa recursiva profunda — captura o custo de cópia de 24 B do CellValue na recursão.
// Profundidades limitadas a 3.000 (a sonda do self-check confirma que não estoura a stack default).
[ShortRunJob]
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class ChainBenchmarks
{
    [Params(1_000, 3_000)]
    public int Depth;

    private INodeObj _obj = null!;
    private INodeObj _boxCache = null!;
    private INodeCv _cellValue = null!;

    [GlobalSetup]
    public void Setup()
    {
        _obj = SpikeTrees.CumChainObj(Depth);
        _boxCache = SpikeTreesBoxCache.CumChain(Depth);
        _cellValue = SpikeTrees.CumChainCv(Depth);
    }

    [Benchmark(Baseline = true), BenchmarkCategory("CumChain")]
    public object? CumChain_Object() => _obj.Eval();

    [Benchmark, BenchmarkCategory("CumChain")]
    public object? CumChain_BoxCache() => _boxCache.Eval();

    [Benchmark, BenchmarkCategory("CumChain")]
    public CellValue CumChain_CellValue() => _cellValue.Eval();
}

// Cenário fiel: grafo de dependências memoizado (ranges + dados mistos + caminhos cruzados), extraído em
// lote. Compara cache Dictionary<string, object?> (Graph_ObjCache, baseline) vs Dictionary<string, CellValue>
// (Graph_CvCache). É AQUI que o boxing mais pesa: no cache object? cada valor numérico é um box de vida longa;
// no cache CellValue o struct fica inline. O cache é reusado (Clear) → mede a alocação por extração.
[ShortRunJob]
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
public class GraphBenchmarks
{
    [Params(10_000, 100_000)]
    public int Cells;

    private ObjEngine _objEngine = null!;
    private CvEngine _cvEngine = null!;

    [GlobalSetup]
    public void Setup()
    {
        var (cells, ids) = GraphBuilder.Build(Cells);
        _objEngine = new ObjEngine(cells, ids);
        _cvEngine = new CvEngine(cells, ids);

        // Aquece o cache uma vez fora da medição (dimensiona os arrays internos do Dictionary).
        _objEngine.ExtractAll();
        _cvEngine.ExtractAll();
    }

    [Benchmark(Baseline = true), BenchmarkCategory("Graph")]
    public double Graph_ObjCache() => _objEngine.ExtractAll();

    [Benchmark, BenchmarkCategory("Graph")]
    public double Graph_CvCache() => _cvEngine.ExtractAll();
}
