using BenchmarkDotNet.Attributes;
using Danfma.MySheet;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Whole-column at scale (plans/whole-column-performance.md). REDUCED shape (50k data cells × 10k
// formulas) so it runs under BenchmarkDotNet as a perf-regression guard. Each iteration invalidates the
// cache and evaluates the whole formula block once — one full pass, mirroring the user's load-once /
// read-once cycle.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --filter '*WholeColumnScale*'
//
// The BigColumn rows are the "consumes the big column" case (still O(F*N) until Phase 2). The
// NarrowColumn rows are the "small columns in a big sheet" case: Phase 1's structural index makes them
// collapse from O(F*N) to O(F * narrow). ShortRunJob keeps the wall-clock sane; Allocated is exact.
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnScaleBenchmarks
{
    private const int DataCells = 50_000;
    private const int Formulas = 10_000;

    [ParamsAllValues]
    public ScaleFormula Formula;

    [ParamsAllValues]
    public ScaleTarget Target;

    private Workbook _workbook = null!;
    private string[] _formulaIds = null!;

    [GlobalSetup]
    public void Setup()
    {
        (_workbook, _formulaIds) = WholeColumnScaleData.Build(DataCells, Formulas, Formula, Target);
    }

    // Drop both caches before every measured pass so each iteration pays the full from-cold cost.
    [IterationSetup]
    public void ResetCaches() => _workbook.InvalidateCache();

    [Benchmark]
    public double EvaluateBlock() => WholeColumnScaleData.EvaluateAll(_workbook, _formulaIds);
}
