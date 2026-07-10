using Danfma.MySheet.Expressions;

namespace Danfma.MySheet.Benchmark.Spike.WholeColumn;

// The data shapes the spike parametrizes over. See plans/whole-column-spike.md for the rationale.
public enum Shape
{
    // A1:A10000 dense, plus B..F with 1.000 cells each. The "good" case for whole-column A.
    DenseColumn,

    // 10.000 cells scattered over 100 columns x random rows (fixed seed). Realistic sparse sheet.
    Sparse,

    // 100.000 cells over 26 columns. Stresses build cost and total footprint.
    Large,

    // Only A1 and A100000. The min/max pathological case that tempts a "bounds scan".
    Pathological,
}

// Builds, once per shape, the three candidate storage layouts side by side so the benchmarks can
// compare them over identical data. Standalone prototype: it never touches the production Sheet.
internal sealed class WholeColumnData
{
    public required Dictionary<string, Expression> StringCells { get; init; }
    public required Dictionary<int, Dictionary<int, Expression>> Tabular { get; init; }
    public required List<(int Col, int Row)> RawCells { get; init; }
    public required List<string> AllIds { get; init; }
    public required int TargetColumn { get; init; }
    public required int TargetRow { get; init; }
    public required string[] LookupIds { get; init; }
    public required (int Col, int Row)[] LookupCoords { get; init; }
    public required int MaxRow { get; init; }

    public static WholeColumnData Build(Shape shape)
    {
        var raw = new List<(int Col, int Row)>();
        var maxRow = 0;

        switch (shape)
        {
            case Shape.DenseColumn:
                for (var r = 1; r <= 10_000; r++)
                {
                    raw.Add((1, r)); // column A
                }

                for (var c = 2; c <= 6; c++) // columns B..F
                {
                    for (var r = 1; r <= 1_000; r++)
                    {
                        raw.Add((c, r));
                    }
                }

                maxRow = 10_000;
                break;

            case Shape.Sparse:
            {
                var rng = new Random(12345); // fixed seed for reproducibility
                var seen = new HashSet<(int, int)>();
                while (seen.Count < 10_000)
                {
                    var c = rng.Next(1, 101); // 100 columns
                    var r = rng.Next(1, 10_001);
                    if (seen.Add((c, r)))
                    {
                        raw.Add((c, r));
                        if (r > maxRow)
                        {
                            maxRow = r;
                        }
                    }
                }

                break;
            }

            case Shape.Large:
            {
                const int perColumn = 100_000 / 26; // 3846
                for (var c = 1; c <= 26; c++)
                {
                    for (var r = 1; r <= perColumn; r++)
                    {
                        raw.Add((c, r));
                    }
                }

                // Top up column A to reach exactly 100.000 cells.
                for (var r = perColumn + 1; raw.Count < 100_000; r++)
                {
                    raw.Add((1, r));
                }

                maxRow = perColumn + (100_000 - 26 * perColumn);
                break;
            }

            case Shape.Pathological:
                raw.Add((1, 1));
                raw.Add((1, 100_000));
                maxRow = 100_000;
                break;
        }

        var stringCells = new Dictionary<string, Expression>(raw.Count);
        var tabular = new Dictionary<int, Dictionary<int, Expression>>();
        var allIds = new List<string>(raw.Count);

        foreach (var (col, row) in raw)
        {
            var id = KeyParse.ToId(col, row);
            Expression value = Expression.Number(row);

            stringCells[id] = value;
            allIds.Add(id);

            if (!tabular.TryGetValue(col, out var colDict))
            {
                tabular[col] = colDict = new Dictionary<int, Expression>();
            }

            colDict[row] = value;
        }

        // Random-order sample of existing cells for the point-lookup hot-path benchmark. Random
        // order defeats any sequential cache-friendliness that would flatter one layout.
        var pool = new List<(int Col, int Row)>(raw);
        var shuffle = new Random(999);
        for (var i = pool.Count - 1; i > 0; i--)
        {
            var j = shuffle.Next(i + 1);
            (pool[i], pool[j]) = (pool[j], pool[i]);
        }

        var sampleSize = Math.Min(1000, pool.Count);
        var lookupCoords = new (int Col, int Row)[sampleSize];
        var lookupIds = new string[sampleSize];
        for (var i = 0; i < sampleSize; i++)
        {
            lookupCoords[i] = pool[i];
            lookupIds[i] = KeyParse.ToId(pool[i].Col, pool[i].Row);
        }

        return new WholeColumnData
        {
            StringCells = stringCells,
            Tabular = tabular,
            RawCells = raw,
            AllIds = allIds,
            TargetColumn = 1, // column A
            TargetRow = 1, // row 1 is populated across the neighbour columns in every shape
            LookupIds = lookupIds,
            LookupCoords = lookupCoords,
            MaxRow = maxRow,
        };
    }

    // The LazyColumnIndex: one pass over the string keys → column -> ids. Rebuilt on invalidation.
    public static Dictionary<int, List<string>> BuildColumnIndex(
        Dictionary<string, Expression> cells
    )
    {
        var index = new Dictionary<int, List<string>>();
        foreach (var id in cells.Keys)
        {
            var col = KeyParse.ColumnOf(id);
            if (!index.TryGetValue(col, out var list))
            {
                index[col] = list = new List<string>();
            }

            list.Add(id);
        }

        return index;
    }
}
