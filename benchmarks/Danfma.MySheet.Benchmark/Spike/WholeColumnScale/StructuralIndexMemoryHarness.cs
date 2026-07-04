using System.Diagnostics;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumnScale;

// Phase-3 (allocation-hygiene 3.3) probe for the NUMERIC structural index.
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --structural-index-memory
//
// Two measurements the plan asked for:
//   (a) the index's retained RAM, string-bucket representation (BEFORE) vs int-bucket representation (AFTER),
//       built from the SAME sheet keys, each isolated by a GC.GetTotalMemory(true) delta. Both share the same
//       already-alive id strings held by the cell dictionary, so the delta is the honest MARGINAL cost of the
//       index structure itself (list backing arrays + dictionary/set overhead), not double-counted strings.
//       A separate line reports the SERIALIZABLE-form ratio (what a persisted index would cost: the id text vs
//       a 4-byte int), which is the ~10x the persistence spike cares about.
//   (b) the open-range hot loop: OLD (one CellAddress parse per populated id, then the dense accessor) vs NEW
//       (the index yields the numeric pair straight into the dense accessor). Same dense hits on both sides, so
//       the delta is exactly the per-id parse the numeric index removes.
public static class StructuralIndexMemoryHarness
{
    private const string DataSheet = "Data";
    private const int Cells = 600_000; // one tall column: ids "A1".."A600000" (~6-7 chars, K1-order magnitude)

    public static void Run()
    {
        Console.WriteLine("== 3.3 numeric structural-index probe ==");
        Console.WriteLine(
            $"Runtime: {Environment.Version}, cores {Environment.ProcessorCount}. "
                + $"Single column A of {Cells:N0} populated cells (ids A1..A{Cells})."
        );
        Console.WriteLine();

        MeasureIndexMemory();
        Console.WriteLine();
        MeasureOpenRangeParse();
    }

    // (a) Retained-RAM of the index structure: real int buckets vs a string-bucket replica of the same keys.
    private static void MeasureIndexMemory()
    {
        var (workbook, sheet) = BuildColumn();

        // AFTER: force the real (int) index — Bucketize over every key + sort the single column bucket. Hold a
        // reference so it cannot be collected before the measurement.
        var baseline = Settle();
        var index = sheet.GetStructuralIndex();
        index.TryGetColumn(1, out var rows); // builds the column map and sorts bucket A
        var afterInt = Settle();
        var intBytes = afterInt - baseline;
        GC.KeepAlive(index);
        GC.KeepAlive(rows);

        // BEFORE: the string-bucket representation the index used to keep — a Dictionary<int, List<string>> of
        // the SAME id strings (shared references, exactly as the old Bucketize stored them; no new strings).
        var beforeBaseline = Settle();
        var stringBuckets = BuildStringBuckets(sheet);
        var afterString = Settle();
        var stringBytes = afterString - beforeBaseline;
        GC.KeepAlive(stringBuckets);

        Console.WriteLine("(a) index retained RAM (marginal, shared id strings excluded — GC delta):");
        Console.WriteLine($"    {"representation",-28} {"retained KB",12} {"bytes/cell",12}");
        Console.WriteLine($"    {"BEFORE List<string> (refs)",-28} {stringBytes / 1024.0,12:N1} {(double)stringBytes / Cells,12:N2}");
        Console.WriteLine($"    {"AFTER  List<int>",-28} {intBytes / 1024.0,12:N1} {(double)intBytes / Cells,12:N2}");
        Console.WriteLine($"    -> in-RAM structure shrinks {(double)stringBytes / intBytes:N2}x (8-byte ref slot -> 4-byte int slot).");

        // SERIALIZABLE form: if the index OWNED its coordinates (the persistence-spike question), the string form
        // pays the id text + object header per cell; the int form pays 4 bytes. This is the ~10x the plan cites.
        long idTextBytes = 0;
        foreach (var id in ColumnIds())
        {
            idTextBytes += ApproxStringBytes(id.Length);
        }

        var intFormBytes = (long)Cells * sizeof(int);
        Console.WriteLine(
            $"    serializable form: id text ~{idTextBytes / 1024.0 / 1024.0:N1} MB vs int "
                + $"{intFormBytes / 1024.0 / 1024.0:N1} MB -> {(double)idTextBytes / intFormBytes:N1}x (persistence-spike lever)."
        );
    }

    // (b) The open-range hot loop: the numeric index removes one CellAddress parse per populated id.
    private static void MeasureOpenRangeParse()
    {
        var (workbook, sheet) = BuildColumn();
        var handle = workbook.ResolveDenseHandle(DataSheet);

        // Prime the dense value cache so both loops measure warm HITS (the common open-range case), isolating the
        // parse from evaluation. A whole-column SUM touches every populated cell of A once.
        var index = sheet.GetStructuralIndex();
        index.TryGetColumn(1, out var rows);

        foreach (var row in rows)
        {
            _ = workbook.GetCellValueDense(handle, DataSheet, 1, row);
        }

        // The string list the OLD index would have held (shared id strings, row-sorted) — built OUTSIDE the timed
        // region so the loop measures only the per-id parse the old ExpandComputedValues paid.
        var ids = new string[rows.Count];
        for (var i = 0; i < rows.Count; i++)
        {
            ids[i] = new CellAddress(1, rows[i]).ToId();
        }

        var (oldMs, oldSum) = TimeBest(() =>
        {
            var acc = 0.0;
            foreach (var id in ids)
            {
                CellAddress.TryGetColumnRow(id, out var column, out var row);
                acc += Numeric(workbook.GetCellValueDense(handle, DataSheet, column, row));
            }

            return acc;
        });

        var (newMs, newSum) = TimeBest(() =>
        {
            var acc = 0.0;
            foreach (var row in rows)
            {
                acc += Numeric(workbook.GetCellValueDense(handle, DataSheet, 1, row));
            }

            return acc;
        });

        Console.WriteLine("(b) open-range hot loop over A:A (warm dense hits), best-of-7 ms:");
        Console.WriteLine($"    {"path",-34} {"ms",10}");
        Console.WriteLine($"    {"OLD  parse id per cell + dense",-34} {oldMs,10:N3}");
        Console.WriteLine($"    {"NEW  numeric pair -> dense",-34} {newMs,10:N3}");
        Console.WriteLine($"    -> the per-id parse died: {(oldMs - newMs) / oldMs * 100:N0}% off the enumeration ({oldMs / Math.Max(newMs, 1e-9):N2}x).");
        Console.WriteLine($"    sums equal: {Math.Abs(oldSum - newSum) < 1e-6} ({oldSum:N0}).");
    }

    private static Dictionary<int, List<string>> BuildStringBuckets(Sheet sheet)
    {
        var buckets = new Dictionary<int, List<string>>();

        foreach (var id in sheet.Keys)
        {
            if (!CellAddress.TryGetColumnRow(id, out var column, out _))
            {
                continue;
            }

            if (!buckets.TryGetValue(column, out var list))
            {
                buckets[column] = list = [];
            }

            list.Add(id);
        }

        return buckets;
    }

    private static (Workbook Workbook, Sheet Sheet) BuildColumn()
    {
        var workbook = new Workbook();
        var sheet = workbook.Sheets.Add(DataSheet);

        for (var r = 1; r <= Cells; r++)
        {
            sheet[$"A{r}"] = new NumberValue(r);
        }

        return (workbook, sheet);
    }

    private static IEnumerable<string> ColumnIds()
    {
        for (var r = 1; r <= Cells; r++)
        {
            yield return new CellAddress(1, r).ToId();
        }
    }

    // A .NET string on 64-bit: object header (16) + length field (4) + 2 bytes/char, rounded up to 8.
    private static long ApproxStringBytes(int length)
    {
        var raw = 16 + 4 + 2L * length;
        return (raw + 7) & ~7L;
    }

    private static double Numeric(ComputedValue value) => value.AsDouble() ?? 0.0;

    private static (double Ms, double Result) TimeBest(Func<double> body)
    {
        var best = double.MaxValue;
        var result = 0.0;

        for (var trial = 0; trial < 7; trial++)
        {
            var stopwatch = Stopwatch.StartNew();
            result = body();
            stopwatch.Stop();
            best = Math.Min(best, stopwatch.Elapsed.TotalMilliseconds);
        }

        return (best, result);
    }

    private static long Settle()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        return GC.GetTotalMemory(forceFullCollection: true);
    }
}
