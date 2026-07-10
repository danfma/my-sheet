using BenchmarkDotNet.Attributes;
using Danfma.MySheet.Excel;

namespace Danfma.MySheet.Benchmark;

// Permanent load-path memory/time gate (Gen0/1/2 + Allocated) for ExcelFile.Load, ahead of the phase G2
// literal-dedup work. The owner's external BDN run on the real ~500k-cell K1 file showed MySheet promoting
// far more into Gen1/Gen2 than Aspose on the same load (77000/43000/8000 vs 99000/30000/5000 Gen0/1/2 over
// 1k ops), ~33% slower overall. This class is the local before/after instrument for that regression: run it
// now, then again once G2 lands, and diff the Gen1/Gen2-per-Gen0 ratios (the promotion signal that matters,
// not the raw Mean).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --filter '*ExcelLoadBenchmarks*'
[MemoryDiagnoser]
public class ExcelLoadBenchmarks
{
    private string? _convertedXlsxPath;
    private string? _syntheticSharedFormulaXlsxPath;

    [GlobalSetup]
    public void Setup()
    {
        var samples = FindSamples();

        // The confidential k1.myxl is preferred when present; the synthetic fixture (built by
        // tools/SyntheticK1Builder) is the committed-workflow fallback with the same K1-like profile.
        var myxl = Path.Combine(samples, "k1.myxl");
        if (!File.Exists(myxl))
        {
            myxl = Path.Combine(samples, "k1-synthetic.myxl");
        }
        if (!File.Exists(myxl))
        {
            throw new FileNotFoundException(
                $"Neither k1.myxl nor k1-synthetic.myxl found in {samples}. "
                    + "Run: dotnet run -c Release --project tools/SyntheticK1Builder"
            );
        }

        Console.WriteLine($"Loading {Path.GetFileName(myxl)} into a MySheet Workbook...");
        var workbook = Danfma.MySheet.Workbook.Load(myxl);

        // Convert to a real .xlsx WITH per-cell formulas (not a values-only flattened snapshot) — the
        // owner's ask: "convert the .myxl to excel and evaluate the load".
        var tmp = Path.Combine(
            Path.GetTempPath(),
            "excel-load-bench-" + Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(tmp);
        _convertedXlsxPath = Path.Combine(tmp, "k1-converted.xlsx");
        workbook.SaveAsExcel(
            _convertedXlsxPath,
            new ExcelExportOptions { FormulaMode = FormulaMode.Formulas }
        );
        Console.WriteLine(
            $"  -> converted to {new FileInfo(_convertedXlsxPath).Length / 1024 / 1024} MB xlsx"
        );

        // Real shared-formula groups (Excel-produced-like file), when the synthetic fixture is present.
        var syntheticXlsx = Path.Combine(samples, "k1-synthetic.xlsx");
        if (File.Exists(syntheticXlsx))
        {
            _syntheticSharedFormulaXlsxPath = syntheticXlsx;
        }
        else
        {
            Console.WriteLine(
                $"(skipping LoadSyntheticSharedFormulas: {syntheticXlsx} not found — "
                    + "run: dotnet run -c Release --project tools/SyntheticK1Builder)"
            );
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (_convertedXlsxPath is not null)
        {
            var dir = Path.GetDirectoryName(_convertedXlsxPath);
            if (dir is not null && Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
    }

    [Benchmark]
    public Danfma.MySheet.Workbook LoadConvertedFromMyxl()
    {
        return ExcelFile.Load(_convertedXlsxPath!);
    }

    [Benchmark]
    public Danfma.MySheet.Workbook? LoadSyntheticSharedFormulas()
    {
        return _syntheticSharedFormulaXlsxPath is null
            ? null
            : ExcelFile.Load(_syntheticSharedFormulaXlsxPath);
    }

    private static string FindSamples()
    {
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 10 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "samples");
            if (Directory.Exists(candidate))
            {
                return candidate;
            }

            dir = Path.GetDirectoryName(dir.TrimEnd(Path.DirectorySeparatorChar));
        }

        return Path.Combine(Directory.GetCurrentDirectory(), "samples");
    }
}
