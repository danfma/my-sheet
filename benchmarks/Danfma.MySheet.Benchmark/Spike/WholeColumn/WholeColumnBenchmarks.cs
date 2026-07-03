using BenchmarkDotNet.Attributes;
using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumn;

// SPIKE: whole-column (A:A) / whole-row (1:1) storage strategy comparison. Pure experiment — no
// production code is touched. Run with:
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --filter '*WholeColumn*'
// ShortRunJob keeps the total time sane; treat Mean deltas under ~15% with suspicion and re-run the
// specific class with a full job. Allocated is measured deterministically even in a short run.

// (1) Whole-column aggregation over column A. NaiveScan (alloc vs no-alloc parse) vs a lazy column
// index (cold build vs warm hit) vs enumerating a tabular column dictionary.
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnScanBenchmarks
{
    [Params(Shape.DenseColumn, Shape.Sparse, Shape.Large)]
    public Shape Shape;

    private WholeColumnData _data = null!;
    private Dictionary<int, List<string>> _prebuiltIndex = null!;

    [GlobalSetup]
    public void Setup()
    {
        _data = WholeColumnData.Build(Shape);
        _prebuiltIndex = WholeColumnData.BuildColumnIndex(_data.StringCells);
    }

    // Baseline: enumerate every key, fully parse it (allocating substring, like CellAddress.Parse),
    // keep the ones in the target column.
    [Benchmark(Baseline = true)]
    public double NaiveScan_Alloc()
    {
        var target = _data.TargetColumn;
        var sum = 0.0;
        foreach (var (id, value) in _data.StringCells)
        {
            var (col, _) = KeyParse.ParseAlloc(id);
            if (col == target)
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }

    // Same scan, no-alloc parse (only the column is needed to filter).
    [Benchmark]
    public double NaiveScan_NoAlloc()
    {
        var target = _data.TargetColumn;
        var sum = 0.0;
        foreach (var (id, value) in _data.StringCells)
        {
            if (KeyParse.ColumnOf(id) == target)
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }

    // Cold: pay the full index construction, then aggregate the target column once.
    [Benchmark]
    public double LazyIndex_Build()
    {
        var index = WholeColumnData.BuildColumnIndex(_data.StringCells);
        return AggregateViaIndex(index);
    }

    // Warm: index already built (reused across an evaluation epoch); just the target-column hit.
    [Benchmark]
    public double LazyIndex_Hit() => AggregateViaIndex(_prebuiltIndex);

    // Tabular: enumerate the target column's inner dictionary directly (no parse at all).
    [Benchmark]
    public double Tabular_ColumnEnum()
    {
        var sum = 0.0;
        if (_data.Tabular.TryGetValue(_data.TargetColumn, out var col))
        {
            foreach (var value in col.Values)
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }

    private double AggregateViaIndex(Dictionary<int, List<string>> index)
    {
        var sum = 0.0;
        if (index.TryGetValue(_data.TargetColumn, out var ids))
        {
            foreach (var id in ids)
            {
                sum += ((NumberValue)_data.StringCells[id]).Value;
            }
        }

        return sum;
    }
}

// (2) Lazy index under cache invalidation. Simulates 100 whole-column reads while the index is
// invalidated (rebuilt) every N reads — N=1 mimics a volatile-heavy sheet, N=100 a stable one.
// Answers: after how many reads per epoch does the index pay for its construction?
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnLazyInvalidationBenchmarks
{
    private const int Reads = 100;

    [Params(1, 10, 100)]
    public int InvalidateEvery;

    private WholeColumnData _data = null!;

    [GlobalSetup]
    public void Setup() => _data = WholeColumnData.Build(Shape.Large);

    // Reference: no index at all, a fresh naive scan per read.
    [Benchmark(Baseline = true)]
    public double NaiveScan_Repeated()
    {
        var target = _data.TargetColumn;
        var total = 0.0;
        for (var read = 0; read < Reads; read++)
        {
            var sum = 0.0;
            foreach (var (id, value) in _data.StringCells)
            {
                if (KeyParse.ColumnOf(id) == target)
                {
                    sum += ((NumberValue)value).Value;
                }
            }

            total += sum;
        }

        return total;
    }

    // Lazy index kept across reads, rebuilt whenever the epoch counter hits InvalidateEvery.
    [Benchmark]
    public double LazyIndex_WithInvalidation()
    {
        var total = 0.0;
        Dictionary<int, List<string>>? index = null;
        var sinceBuild = 0;

        for (var read = 0; read < Reads; read++)
        {
            if (index is null || sinceBuild >= InvalidateEvery)
            {
                index = WholeColumnData.BuildColumnIndex(_data.StringCells);
                sinceBuild = 0;
            }

            var sum = 0.0;
            if (index.TryGetValue(_data.TargetColumn, out var ids))
            {
                foreach (var id in ids)
                {
                    sum += ((NumberValue)_data.StringCells[id]).Value;
                }
            }

            total += sum;
            sinceBuild++;
        }

        return total;
    }
}

// (3) THE hot path: point lookup of a single populated cell, 1.000 random existing ids per op.
// StringDict[id] (today) vs Tabular[col][row]. This is the regression that would tax EVERY
// evaluation, so it is the most important measurement in the spike.
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnPointLookupBenchmarks
{
    [Params(Shape.DenseColumn, Shape.Sparse, Shape.Large)]
    public Shape Shape;

    private WholeColumnData _data = null!;

    [GlobalSetup]
    public void Setup() => _data = WholeColumnData.Build(Shape);

    // Baseline: the current storage. One dictionary probe on the string id.
    [Benchmark(Baseline = true)]
    public double PointLookup_StringDict()
    {
        var sum = 0.0;
        foreach (var id in _data.LookupIds)
        {
            sum += ((NumberValue)_data.StringCells[id]).Value;
        }

        return sum;
    }

    // Tabular with the id arriving as a string (must be parsed): span-parse + two dictionary probes.
    [Benchmark]
    public double PointLookup_Tabular_ParseSpan()
    {
        var sum = 0.0;
        foreach (var id in _data.LookupIds)
        {
            var (col, row) = KeyParse.ParseSpan(id);
            sum += ((NumberValue)_data.Tabular[col][row]).Value;
        }

        return sum;
    }

    // Tabular best case: coordinates already known (references pre-parsed to col/row). Isolates the
    // pure double-dictionary cost, without any parse.
    [Benchmark]
    public double PointLookup_Tabular_PreParsed()
    {
        var sum = 0.0;
        foreach (var (col, row) in _data.LookupCoords)
        {
            sum += ((NumberValue)_data.Tabular[col][row]).Value;
        }

        return sum;
    }
}

// (4) The symmetric bad case for Tabular: whole-ROW (1:1) aggregation. A string scan filters keys by
// row exactly like it filters by column; the column-major tabular layout must probe every column.
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnRowEnumBenchmarks
{
    [Params(Shape.DenseColumn, Shape.Sparse, Shape.Large)]
    public Shape Shape;

    private WholeColumnData _data = null!;

    [GlobalSetup]
    public void Setup() => _data = WholeColumnData.Build(Shape);

    [Benchmark(Baseline = true)]
    public double RowScan_StringDict()
    {
        var target = _data.TargetRow;
        var sum = 0.0;
        foreach (var (id, value) in _data.StringCells)
        {
            if (KeyParse.RowOf(id) == target)
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }

    // Column-major storage → whole-row means probing every column dictionary. Gets worse the more
    // columns exist.
    [Benchmark]
    public double RowEnum_Tabular()
    {
        var target = _data.TargetRow;
        var sum = 0.0;
        foreach (var col in _data.Tabular.Values)
        {
            if (col.TryGetValue(target, out var value))
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }
}

// (5) Pathological "min/max" case: only A1 and A100000 populated. Documents why a bounds-based range
// scan (iterate A1..A100000) is catastrophic on sparse data, versus a scan over populated cells.
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnBoundsBenchmarks
{
    private WholeColumnData _data = null!;

    [GlobalSetup]
    public void Setup() => _data = WholeColumnData.Build(Shape.Pathological);

    // The "populated cells" semantics: two keys, done.
    [Benchmark(Baseline = true)]
    public double NaiveScan_PopulatedOnly()
    {
        var target = _data.TargetColumn;
        var sum = 0.0;
        foreach (var (id, value) in _data.StringCells)
        {
            if (KeyParse.ColumnOf(id) == target)
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }

    // The "bounds scan" temptation: walk every row from 1 to MaxRow probing the string dictionary.
    // 100.000 TryGetValue + 100.000 id allocations to find 2 cells.
    [Benchmark]
    public double BoundsRangeScan_StringDict()
    {
        var sum = 0.0;
        for (var row = 1; row <= _data.MaxRow; row++)
        {
            var id = KeyParse.ToId(_data.TargetColumn, row);
            if (_data.StringCells.TryGetValue(id, out var value))
            {
                sum += ((NumberValue)value).Value;
            }
        }

        return sum;
    }

    // Same bounds walk against the tabular column dictionary. No id allocation, but still 100.000
    // probes to find 2 cells — bounds scanning is wrong regardless of layout.
    [Benchmark]
    public double BoundsRangeScan_Tabular()
    {
        var sum = 0.0;
        if (_data.Tabular.TryGetValue(_data.TargetColumn, out var col))
        {
            for (var row = 1; row <= _data.MaxRow; row++)
            {
                if (col.TryGetValue(row, out var value))
                {
                    sum += ((NumberValue)value).Value;
                }
            }
        }

        return sum;
    }
}

// (6) Total footprint proxy: allocations to construct each layout from the raw cell list. Reported
// via the Allocated column of MemoryDiagnoser.
[ShortRunJob]
[MemoryDiagnoser]
public class WholeColumnMemoryBenchmarks
{
    [Params(Shape.DenseColumn, Shape.Sparse, Shape.Large)]
    public Shape Shape;

    private WholeColumnData _data = null!;

    [GlobalSetup]
    public void Setup() => _data = WholeColumnData.Build(Shape);

    [Benchmark(Baseline = true)]
    public object Build_StringDict()
    {
        var cells = new Dictionary<string, Expression>(_data.RawCells.Count);
        foreach (var (col, row) in _data.RawCells)
        {
            cells[KeyParse.ToId(col, row)] = Expression.Number(row);
        }

        return cells;
    }

    [Benchmark]
    public object Build_Tabular()
    {
        var tabular = new Dictionary<int, Dictionary<int, Expression>>();
        foreach (var (col, row) in _data.RawCells)
        {
            if (!tabular.TryGetValue(col, out var colDict))
            {
                tabular[col] = colDict = new Dictionary<int, Expression>();
            }

            colDict[row] = Expression.Number(row);
        }

        return tabular;
    }

    [Benchmark]
    public object Build_LazyIndexOnTopOfStringDict()
    {
        // The lazy index is an ADDITION to the existing string dictionary, not a replacement.
        return WholeColumnData.BuildColumnIndex(_data.StringCells);
    }
}
