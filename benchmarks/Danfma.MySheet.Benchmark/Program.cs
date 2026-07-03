using BenchmarkDotNet.Running;
using Danfma.MySheet.Benchmark;
using Danfma.MySheet.Benchmark.Spike;
using Danfma.MySheet.Benchmark.Spike.MessagePackFormat;
using Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Sanidade do spike CellValue (Fase 1): `dotnet run -- --check`.
if (args.Contains("--check"))
{
    SpikeSelfCheck.Run();
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

// Spike MessagePack format — byte sizes: `dotnet run -c Release -- --messagepack-size`.
if (args.Contains("--messagepack-size"))
{
    MessagePackSizeReport.Run();
    return;
}

// Spike CellValue (Fase 2): `dotnet run -c Release -- --filter *Spike*`.
// Sem filtro cai no menu do BenchmarkSwitcher; SheetBenchmarks continua disponível.
BenchmarkSwitcher.FromAssembly(typeof(SheetBenchmarks).Assembly).Run(args);
