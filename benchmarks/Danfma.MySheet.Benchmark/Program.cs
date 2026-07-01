using BenchmarkDotNet.Running;
using Danfma.MySheet.Benchmark;
using Danfma.MySheet.Benchmark.Spike;

// Sanidade do spike CellValue (Fase 1): `dotnet run -- --check`.
if (args.Contains("--check"))
{
    SpikeSelfCheck.Run();
    return;
}

// Spike CellValue (Fase 2): `dotnet run -c Release -- --filter *Spike*`.
// Sem filtro cai no menu do BenchmarkSwitcher; SheetBenchmarks continua disponível.
BenchmarkSwitcher.FromAssembly(typeof(SheetBenchmarks).Assembly).Run(args);
