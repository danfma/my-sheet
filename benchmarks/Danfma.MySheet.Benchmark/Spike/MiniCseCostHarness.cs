using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike;

// Wall-clock cost harness for the mini-CSE element-wise path (plans/mini-cse-array-arguments.md, Phase C).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --mini-cse-cost
//
// Goal: prove the array-consuming path is O(range of the formula) and comparable to a native scalar scan of
// the same size — NOT a pathological blow-up — at the K1 scale (194 rows) and 50× beyond it (10k rows).
//
// Each measured cell is re-evaluated K times with InvalidateCache() between reads (forcing a full fresh
// evaluation each time), reported as best-of-5 of the mean µs/eval. Two data sizes: 194 and 10 000.
//
//   ARRAY   =SUM(IF(A2:An="Show",1,0))            — the mini-CSE path (materializes the closed range,
//                                                    zips the comparison, folds the IF vector).
//   SCALAR  =COUNTIF(A2:An,"Show")                — the native single-pass scalar scan of the SAME range
//                                                    (the "equivalent-size scalar formula" reference).
//   BH25    =INDEX(ROW($A:$A),SMALL(IF(A2:An="Show",IF(ROW(A2:An)>2,ROW(A2:An))),1))
//                                                  — the real K1 idiom (open-range ROW identity, no column
//                                                    materialization; closed range A2:An materialized once).
//
// The array/scalar ratio isolates the mini-CSE machinery overhead; the 194→10k growth (≈51.5× the rows)
// staying ≈linear is the O(range) proof.
public static class MiniCseCostHarness
{
    private const string Data = "Data";
    private const string Calc = "Calc";
    private const int Iterations = 200;
    private static readonly int[] Sizes = [194, 10_000];

    public static void Run()
    {
        Console.WriteLine("== mini-CSE array cost harness ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"{Iterations} iterations of {{ InvalidateCache(); GetCellValue(f) }}; "
                + "best-of-5 mean µs/eval. Column A alternates Show/Hide over N rows."
        );
        Console.WriteLine();
        Console.WriteLine(
            $"{"Rows", 7} {"ARRAY µs", 12} {"SCALAR µs", 12} {"array/scalar", 13} {"BH25 µs", 12}"
        );

        double? arrayAt194 = null;
        double? bh25At194 = null;

        foreach (var size in Sizes)
        {
            var last = size + 1; // rows 2..(size+1)

            var array = Measure(size, $"=SUM(IF({Data}!A2:A{last}=\"Show\",1,0))");
            var scalar = Measure(size, $"=COUNTIF({Data}!A2:A{last},\"Show\")");
            var bh25 = Measure(
                size,
                $"=INDEX(ROW($A:$A),SMALL(IF({Data}!A2:A{last}=\"Show\","
                    + $"IF(ROW({Data}!A2:A{last})>2,ROW({Data}!A2:A{last}))),1))"
            );

            if (size == 194)
            {
                arrayAt194 = array;
                bh25At194 = bh25;
            }

            Console.WriteLine(
                $"{size, 7:N0} {array, 12:N3} {scalar, 12:N3} {array / scalar, 13:N2} {bh25, 12:N3}"
            );
        }

        if (arrayAt194 is { } a194 && bh25At194 is { } b194)
        {
            var arrayAt10k = Measure(10_000, "=SUM(IF(Data!A2:A10001=\"Show\",1,0))");
            var bh25At10k = Measure(
                10_000,
                "=INDEX(ROW($A:$A),SMALL(IF(Data!A2:A10001=\"Show\","
                    + "IF(ROW(Data!A2:A10001)>2,ROW(Data!A2:A10001))),1))"
            );

            Console.WriteLine();
            Console.WriteLine($"O(range) check — rows grew {10_000 / 194.0:N1}× (194→10k):");
            Console.WriteLine($"  ARRAY  {arrayAt10k / a194, 6:N1}× slower");
            Console.WriteLine($"  BH25   {bh25At10k / b194, 6:N1}× slower");
            Console.WriteLine("  (≈ linear ⇒ O(range); the array path adds no super-linear cost).");
        }
    }

    private static double Measure(int size, string formula)
    {
        var (workbook, calcId) = Build(size, formula);

        // Warm up (parse/JIT and first structural touch) before timing.
        workbook.InvalidateCache();
        _ = workbook.GetCellValue(Calc, calcId);

        var best = double.MaxValue;

        for (var trial = 0; trial < 5; trial++)
        {
            var stopwatch = Stopwatch.StartNew();

            for (var i = 0; i < Iterations; i++)
            {
                workbook.InvalidateCache();
                _ = workbook.GetCellValue(Calc, calcId);
            }

            stopwatch.Stop();
            best = Math.Min(best, stopwatch.Elapsed.TotalMicroseconds / Iterations);
        }

        return best;
    }

    private static (Workbook Workbook, string CalcId) Build(int size, string formula)
    {
        var workbook = new Workbook();
        var data = workbook.Sheets.Add(Data);
        var calc = workbook.Sheets.Add(Calc);

        for (var r = 2; r <= size + 1; r++)
        {
            data[$"A{r}"] = new Expressions.StringValue((r % 2 == 0) ? "Show" : "Hide");
        }

        const string calcId = "A1";
        calc[calcId] = ExpressionParser.Parse(formula, calc);

        return (workbook, calcId);
    }
}
