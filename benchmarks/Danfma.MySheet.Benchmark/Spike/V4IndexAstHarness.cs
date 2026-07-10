using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

namespace Danfma.MySheet.Benchmark.Spike;

// Spike v4 (pós-3.3): probes for the two owner questions.
//   Theme 1 — persistence of the structural index.
//   Theme 2 — revisit of the numeric AST (the deferred breaking).
//
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --v4-index-rebuild
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --v4-index-persist
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --v4-resident
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --v4-parse
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --v4-hotpath
//   dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --v4-all
//
// METHODOLOGY: single process, best-of-N (min) for wall-clock; retained memory via GC.GetTotalMemory(true)
// deltas around a settle; churn via GC.GetTotalAllocatedBytes(true). N reported per table. Nothing here
// touches production — every probe lives in the benchmark project and reaches internals through the existing
// InternalsVisibleTo. The numeric-AST probes model the string surfaces a numeric AST WOULD remove; they do
// not build a full numeric AST (that is the breaking change under evaluation).
public static class V4IndexAstHarness
{
    private const int Best = 7;

    public static void RunAll(string[] args)
    {
        IndexRebuild();
        Console.WriteLine();
        IndexPersist();
        Console.WriteLine();
        Resident();
        Console.WriteLine();
        Parse();
        Console.WriteLine();
        HotPath();
    }

    // =====================================================================================================
    // THEME 1.1 — the rebuild ceiling: what the persisted index would REPLACE.
    // =====================================================================================================
    // The structural index is runtime-only and rebuilt lazily on the first open-range read after a Load. Its
    // build is one O(sheet) Bucketize pass + a per-bucket sort on first access. This isolates that pure cost
    // (no formula evaluation mixed in) across 100k/600k/1M cells, in the realistic K1 magnitude, and also the
    // end-to-end "Load a cold workbook then first open-range touch" so the ceiling is framed honestly.
    public static void IndexRebuild()
    {
        Console.WriteLine(
            "== THEME 1.1 — structural-index rebuild ceiling (pure build, no eval) =="
        );
        Console.WriteLine(
            $"Runtime: {Environment.Version}. Sheet = N cells spread over {ColumnsForShape} columns "
                + "(realistic multi-column open-range target). Best-of-"
                + Best
                + " min ms."
        );
        Console.WriteLine(
            $"{"Cells", 12} {"Bucketize+sortAll ms", 22} {"col-map only ms", 18} {"bytes/cell (retained)", 22}"
        );

        foreach (var n in new[] { 100_000, 600_000, 1_000_000 })
        {
            // Fresh sheet per trial (the index build is once-per-life, so we cannot re-time on the same sheet).
            var best = double.MaxValue;
            var bestColOnly = double.MaxValue;

            for (var t = 0; t < Best; t++)
            {
                var sheet = BuildSpreadSheet(n, out _);
                var sw = Stopwatch.StartNew();
                var index = sheet.GetStructuralIndex();
                // Force the column map build (Bucketize) + sort EVERY column bucket (first-access sort).
                foreach (var col in index.ColumnKeys)
                {
                    index.TryGetColumn(col, out _);
                }
                sw.Stop();
                best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
                GC.KeepAlive(index);

                var sheet2 = BuildSpreadSheet(n, out _);
                var sw2 = Stopwatch.StartNew();
                var index2 = sheet2.GetStructuralIndex();
                _ = index2.ColumnKeys.Count; // Bucketize only, no per-bucket sort
                sw2.Stop();
                bestColOnly = Math.Min(bestColOnly, sw2.Elapsed.TotalMilliseconds);
                GC.KeepAlive(index2);
            }

            // Retained bytes of the built index (marginal, shared id strings excluded).
            var sheetR = BuildSpreadSheet(n, out _);
            var baseR = Settle();
            var idx = sheetR.GetStructuralIndex();
            foreach (var col in idx.ColumnKeys)
            {
                idx.TryGetColumn(col, out _);
            }
            var afterR = Settle();
            GC.KeepAlive(idx);
            var perCell = (double)(afterR - baseR) / n;

            Console.WriteLine($"{n, 12:N0} {best, 22:N2} {bestColOnly, 18:N2} {perCell, 22:N2}");
        }
    }

    // =====================================================================================================
    // THEME 1.2 — persistence prototype of the numeric index form.
    // =====================================================================================================
    // Serialize/deserialize the numeric index (per sheet: column -> ascending int rows). Measure: bytes raw
    // and Brotli (in the spirit of MSWM v2), the extra save cost, and the load+validate cost that would
    // REPLACE the O(sheet) rebuild. Validation = cell-count check against Sheet.Count (a cheap consistency
    // gate; a self-healing section simply rebuilds on mismatch).
    public static void IndexPersist()
    {
        Console.WriteLine("== THEME 1.2 — index persistence prototype (numeric form) ==");
        Console.WriteLine(
            "Per sheet: column -> ascending int rows. Encoding: [colCount][ per col: col int32, "
                + "rowCount int32, rows int32[] ]. Brotli(Optimal) over the whole blob. Best-of-"
                + Best
                + "."
        );
        Console.WriteLine(
            $"{"Cells", 12} {"raw KB", 10} {"brotli KB", 11} {"serialize ms", 13} {"deser+valid ms", 15} {"rebuild ms", 12}"
        );

        foreach (var n in new[] { 100_000, 600_000, 1_000_000 })
        {
            var sheet = BuildSpreadSheet(n, out var totalCells);
            var index = sheet.GetStructuralIndex();
            var columns = new List<(int Col, int[] Rows)>();
            foreach (var col in index.ColumnKeys.OrderBy(c => c))
            {
                index.TryGetColumn(col, out var rows);
                columns.Add((col, rows.ToArray()));
            }

            // Serialize (best-of-N min).
            byte[] raw = null!;
            var serMs = double.MaxValue;
            for (var t = 0; t < Best; t++)
            {
                var sw = Stopwatch.StartNew();
                raw = SerializeIndex(columns);
                sw.Stop();
                serMs = Math.Min(serMs, sw.Elapsed.TotalMilliseconds);
            }

            var brotli = BrotliBytes(raw);

            // Deserialize + validate (best-of-N min). Validation: reconstruct the Dictionary<int,List<int>>
            // and check the summed cell count against the sheet's real Count (the consistency gate).
            var deserMs = double.MaxValue;
            var validatedCount = 0;
            for (var t = 0; t < Best; t++)
            {
                var sw = Stopwatch.StartNew();
                var rebuilt = DeserializeIndex(raw, out var count);
                var ok = count == totalCells;
                sw.Stop();
                if (!ok)
                {
                    throw new InvalidOperationException(
                        $"validation failed: {count} != {totalCells}"
                    );
                }
                validatedCount = count;
                deserMs = Math.Min(deserMs, sw.Elapsed.TotalMilliseconds);
                GC.KeepAlive(rebuilt);
            }

            // The rebuild it replaces (Bucketize + sort all), best-of-N on fresh sheets.
            var rebuildMs = double.MaxValue;
            for (var t = 0; t < Best; t++)
            {
                var s = BuildSpreadSheet(n, out _);
                var sw = Stopwatch.StartNew();
                var idx = s.GetStructuralIndex();
                foreach (var col in idx.ColumnKeys)
                {
                    idx.TryGetColumn(col, out _);
                }
                sw.Stop();
                rebuildMs = Math.Min(rebuildMs, sw.Elapsed.TotalMilliseconds);
                GC.KeepAlive(idx);
            }

            Console.WriteLine(
                $"{n, 12:N0} {raw.Length / 1024.0, 10:N1} {brotli.Length / 1024.0, 11:N1} "
                    + $"{serMs, 13:N2} {deserMs, 15:N2} {rebuildMs, 12:N2}"
            );
            _ = validatedCount;
        }
    }

    // =====================================================================================================
    // THEME 2.1 — resident string surfaces of the K1 model.
    // =====================================================================================================
    // Build the real K1-shape model (400k formulas cross-sheet) and measure its retained heap, then attribute
    // the string families a numeric AST would remove, EACH measured as an isolated retained-heap delta:
    //   (1) _cells keys — the string dictionary keys (NON-BREAKING to numericize: internal dictionary).
    //   (2) CellReference.Id + SheetName strings inside the AST nodes (BREAKING: wire format).
    // The lever comparison (1) vs (2) is the crux of the "minimum scope" verdict.
    public static void Resident()
    {
        Console.WriteLine("== THEME 2.1 — resident string surfaces (K1 model, 400k formulas) ==");

        // Full model retained (ground truth).
        var baseAll = Settle();
        var wb = BuildK1Model(out var dataCells, out var formulaCells, out var refNodes);
        var afterAll = Settle();
        var totalRetained = afterAll - baseAll;
        GC.KeepAlive(wb);

        Console.WriteLine(
            $"Data cells {dataCells:N0}, formula cells {formulaCells:N0}, CellReference nodes {refNodes:N0}. "
                + $"Full model retained: {totalRetained / 1024.0 / 1024.0:N1} MB."
        );
        Console.WriteLine();

        // (1) _cells key strings: collect every id key of every sheet into an array; measure retained. These
        // are the exact string objects a Dictionary<(int,int),Expression> would drop (the entries stay, only
        // the key representation changes from an 8-byte ref + string object to an 8-byte struct in-place).
        var keys = new List<string>(dataCells + formulaCells);
        foreach (var (_, sheet) in wb.Sheets)
        {
            foreach (var id in sheet.Keys)
            {
                keys.Add(id);
            }
        }

        long cellKeyBytes = 0;
        foreach (var k in keys)
        {
            cellKeyBytes += ApproxStringBytes(k.Length);
        }

        // (2) AST-node strings: the Id + SheetName held by every CellReference. We re-derive the exact same
        // strings the parser produced (Id via ToId, SheetName the qualifier) and size them. SheetName is a
        // FRESH allocation per node in the current parser (the tokenizer substring), so it counts per node.
        long astIdBytes = 0;
        long astSheetBytes = 0;
        var refIds = CollectCellReferenceStrings(
            wb,
            out astIdBytes,
            out astSheetBytes,
            out var distinctSheetNames
        );

        Console.WriteLine(
            "string family (theoretical bytes, object-header model 16+4+2·len, 8-aligned):"
        );
        Console.WriteLine($"    {"family", -38} {"count", 12} {"MB", 10} {"note", -18}");
        Console.WriteLine(
            $"    {"(1) _cells keys", -38} {keys.Count, 12:N0} {cellKeyBytes / 1024.0 / 1024.0, 10:N1} {"NON-BREAKING", -18}"
        );
        Console.WriteLine(
            $"    {"(2a) CellReference.Id", -38} {refIds, 12:N0} {astIdBytes / 1024.0 / 1024.0, 10:N1} {"breaking (wire)", -18}"
        );
        Console.WriteLine(
            $"    {"(2b) CellReference.SheetName (per node)", -38} {refIds, 12:N0} {astSheetBytes / 1024.0 / 1024.0, 10:N1} {"breaking (wire)", -18}"
        );
        Console.WriteLine(
            $"    -> NON-BREAKING lever (_cells keys): {cellKeyBytes / 1024.0 / 1024.0:N1} MB; "
                + $"BREAKING lever (AST Id+SheetName): {(astIdBytes + astSheetBytes) / 1024.0 / 1024.0:N1} MB."
        );
        Console.WriteLine(
            $"    (distinct sheet-name strings actually alive if interned: {distinctSheetNames}.)"
        );

        var cellCount = keys.Count;
        keys = null!; // release so the dictionaries below are the SOLE owners of their key strings.
        GC.KeepAlive(wb);
        wb = null!; // release the model too: the key-representation probe must own its strings alone.

        // Empirical isolation of the _cells key representation: same values, string key vs (int,int) key.
        // The string dict GENERATES its keys inline (sole owner), so freeing them in the numeric case is a
        // real delta — unlike a probe that kept a separate keys list alive (which would hide the string cost).
        MeasureCellsKeyRepresentation(cellCount);
    }

    // Build a Dictionary<string,Expr> whose keys it SOLELY owns (generated inline) vs a Dictionary<(int,int),
    // Expr>, measuring each retained separately. The delta is exactly the id-string objects the numeric key
    // drops — the non-breaking _cells lever, measured rather than modelled. Keys mirror the K1 magnitude
    // (7-char ids "C123456").
    private static void MeasureCellsKeyRepresentation(int cellCount)
    {
        var shared = BlankValue.Instance; // one shared Expression: isolates the key/entry cost, not values.

        var baseS = Settle();
        var stringDict = new Dictionary<string, Expression>(cellCount);
        for (var i = 0; i < cellCount; i++)
        {
            stringDict["C" + i] = shared; // fresh string, owned only by the dictionary
        }
        var afterS = Settle();
        var stringBytes = afterS - baseS;
        GC.KeepAlive(stringDict);
        stringDict = null!;

        var baseI = Settle();
        var intDict = new Dictionary<(int, int), Expression>(cellCount);
        for (var i = 0; i < cellCount; i++)
        {
            intDict[(3, i)] = shared;
        }
        var afterI = Settle();
        var intBytes = afterI - baseI;
        GC.KeepAlive(intDict);

        Console.WriteLine();
        Console.WriteLine(
            "_cells key representation, MEASURED retained (dict is sole key owner, entries included):"
        );
        Console.WriteLine(
            $"    Dictionary<string,Expr>   {stringBytes / 1024.0 / 1024.0, 8:N1} MB  ({(double)stringBytes / cellCount:N1} B/cell)"
        );
        Console.WriteLine(
            $"    Dictionary<(int,int),Expr>{intBytes / 1024.0 / 1024.0, 8:N1} MB  ({(double)intBytes / cellCount:N1} B/cell)"
        );
        Console.WriteLine(
            $"    -> non-breaking key swap saves {(stringBytes - intBytes) / 1024.0 / 1024.0:N1} MB ({(double)stringBytes / Math.Max(intBytes, 1):N2}x)."
        );
    }

    // =====================================================================================================
    // THEME 2.2 — parse churn.
    // =====================================================================================================
    // How much of the parse allocation is the id/sheet string creation that a numeric AST would avoid?
    // Measured: total parse churn of the 400k K1 formulas; and, in isolation, the churn of the string ops
    // the parser runs per reference (NormalizeCellId's ToUpperInvariant + the tokenizer sheet substring).
    public static void Parse()
    {
        Console.WriteLine("== THEME 2.2 — parse churn (400k K1 formulas) ==");

        var formulas = BuildK1FormulaStrings(out var count);
        var sheet = new Sheet { Name = "S" };

        // Total parse churn (best-of-N min).
        var totalChurn = long.MaxValue;
        var totalMs = double.MaxValue;
        for (var t = 0; t < 3; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            Expression sink = null!;
            foreach (var f in formulas)
            {
                sink = ExpressionParser.Parse(f, sheet);
            }
            sw.Stop();
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            totalChurn = Math.Min(totalChurn, churn);
            totalMs = Math.Min(totalMs, sw.Elapsed.TotalMilliseconds);
            GC.KeepAlive(sink);
        }

        Console.WriteLine(
            $"formulas: {count:N0}. Total parse: {totalMs:N0} ms, churn {totalChurn / 1024.0 / 1024.0:N1} MB ({(double)totalChurn / count:N0} B/formula)."
        );

        // Isolate the per-reference string ops. Each K1 formula holds refs like Data!A{v}, Data!B{v}. Count
        // the reference tokens and measure the churn of just the string work the parser does per reference:
        //   - NormalizeCellId: StripDollars + ToUpperInvariant (ToUpperInvariant always allocates a new string)
        //   - the SheetName substring the tokenizer hands to the qualified reference
        var refTokens = CountReferenceStrings(out var refCount);
        var strChurn = long.MaxValue;
        for (var t = 0; t < 3; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            string sink = null!;
            foreach (var (id, sheetName) in refTokens)
            {
                sink = id.ToUpperInvariant(); // NormalizeCellId's allocation
                sink = new string(sheetName); // the fresh SheetName the tokenizer substring produces
            }
            var churn = GC.GetTotalAllocatedBytes(true) - before;
            strChurn = Math.Min(strChurn, churn);
            GC.KeepAlive(sink);
        }

        Console.WriteLine(
            $"reference nodes: {refCount:N0}. Id/sheet string ops churn (isolated): "
                + $"{strChurn / 1024.0 / 1024.0:N1} MB = {(double)strChurn / Math.Max(totalChurn, 1) * 100:N0}% of total parse churn."
        );
        Console.WriteLine(
            "NOTE: warm-start (persisted values) skips the ENTIRE parse+compute on Load; this churn only "
                + "applies to a cold Load or an initial build — the numeric AST competes with warm-start here."
        );
    }

    // =====================================================================================================
    // THEME 2.3 — remaining hot-path string touches.
    // =====================================================================================================
    // Two residual touches a numeric AST/int-keyed store would remove:
    //   (A) SetCell is string-keyed: populating 600k cells hashes strings. Compare to an (int,int) key.
    //   (B) GetCellValueDense MISS materializes the A1 id via CellAddress.ToId (a StringBuilder allocation).
    //       Measure the aggregate ToId cost at K1 miss volume.
    public static void HotPath()
    {
        Console.WriteLine("== THEME 2.3 — residual hot-path string touches ==");

        const int n = 600_000;
        var ids = new string[n];
        var coords = new (int, int)[n];
        for (var i = 0; i < n; i++)
        {
            var col = (i % 4) + 1;
            var row = (i / 4) + 1;
            coords[i] = (col, row);
            ids[i] = new CellAddress(col, row).ToId();
        }

        // (A) SetCell-shape: string-keyed dictionary insert vs (int,int)-keyed.
        var shared = BlankValue.Instance;
        var strMs = double.MaxValue;
        var strChurn = long.MaxValue;
        for (var t = 0; t < Best; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            var d = new Dictionary<string, Expression>(n);
            for (var i = 0; i < n; i++)
            {
                d[ids[i]] = shared;
            }
            sw.Stop();
            strMs = Math.Min(strMs, sw.Elapsed.TotalMilliseconds);
            strChurn = Math.Min(strChurn, GC.GetTotalAllocatedBytes(true) - before);
            GC.KeepAlive(d);
        }

        var intMs = double.MaxValue;
        for (var t = 0; t < Best; t++)
        {
            GcQuiesce();
            var sw = Stopwatch.StartNew();
            var d = new Dictionary<(int, int), Expression>(n);
            for (var i = 0; i < n; i++)
            {
                d[coords[i]] = shared;
            }
            sw.Stop();
            intMs = Math.Min(intMs, sw.Elapsed.TotalMilliseconds);
            GC.KeepAlive(d);
        }

        Console.WriteLine($"(A) populate {n:N0} cells (string keys already alive):");
        Console.WriteLine(
            $"    string-keyed insert {strMs, 8:N1} ms  (churn {strChurn / 1024.0 / 1024.0:N1} MB)"
        );
        Console.WriteLine($"    (int,int)-keyed     {intMs, 8:N1} ms");

        // (B) ToId on a miss: the aggregate cost if every one of the K1 formula cells took a dense miss (the
        // first compute of a range-expanded reference). Measure ToId over n coords.
        var toIdMs = double.MaxValue;
        var toIdChurn = long.MaxValue;
        for (var t = 0; t < Best; t++)
        {
            GcQuiesce();
            var before = GC.GetTotalAllocatedBytes(true);
            var sw = Stopwatch.StartNew();
            string sink = null!;
            for (var i = 0; i < n; i++)
            {
                sink = new CellAddress(coords[i].Item1, coords[i].Item2).ToId();
            }
            sw.Stop();
            toIdMs = Math.Min(toIdMs, sw.Elapsed.TotalMilliseconds);
            toIdChurn = Math.Min(toIdChurn, GC.GetTotalAllocatedBytes(true) - before);
            GC.KeepAlive(sink);
        }

        Console.WriteLine(
            $"(B) CellAddress.ToId over {n:N0} misses: {toIdMs, 6:N1} ms, churn {toIdChurn / 1024.0 / 1024.0:N1} MB."
        );
        Console.WriteLine(
            "    NOTE: single-cell references (Data!B{v}) take the no-alloc TryGetColumnRow path in "
                + "GetCellValue — ToId only fires on a RANGE-expansion miss, once per cell per epoch."
        );
    }

    // ---------------------------------------------------------------------------------- shape builders

    private const int ColumnsForShape = 4;

    // A sheet of n cells spread over ColumnsForShape columns (rows 1..n/cols) — a realistic open-range target
    // with several buckets. Returns the Sheet and the exact total cell count.
    private static Sheet BuildSpreadSheet(int n, out int totalCells)
    {
        var wb = new Workbook();
        var sheet = wb.Sheets.Add("Data");
        var perCol = n / ColumnsForShape;
        var count = 0;
        for (var col = 1; col <= ColumnsForShape; col++)
        {
            var name = ColumnName(col);
            for (var r = 1; r <= perCol; r++)
            {
                sheet[$"{name}{r}"] = NumberNode;
                count++;
            }
        }

        totalCells = count;
        return sheet;
    }

    private static readonly Expression NumberNode = new NumberValue(1);

    // The real K1 model: Data 100k×2, S 400k formulas cross-sheet (same shape as K1EndToEndHarness).
    private static Workbook BuildK1Model(out int dataCells, out int formulaCells, out int refNodes)
    {
        const int dataRows = 100_000;
        const int formulaRows = 200_000;
        const int showEvery = 7;
        var wb = new Workbook();
        var data = wb.Sheets.Add("Data");
        var s = wb.Sheets.Add("S");
        var show = new StringValue("Show");
        var hide = new StringValue("Hide");

        for (var v = 1; v <= dataRows; v++)
        {
            data["A" + v] = v % showEvery == 1 ? show : hide;
            data["B" + v] = new NumberValue(v);
        }

        var refs = 0;
        for (var r = 1; r <= formulaRows; r++)
        {
            var v = (r % dataRows) + 1;
            s["C" + r] = ExpressionParser.Parse(
                $"=IF(Data!A{v}=\"Show\",Data!B{v}*2,Data!B{v}/2)",
                s
            );
            s["D" + r] = ExpressionParser.Parse($"=Data!B{v}*2+1", s);
            refs += 4; // C has 3 CellReferences, D has 1
        }

        dataCells = dataRows * 2;
        formulaCells = formulaRows * 2;
        refNodes = refs;
        return wb;
    }

    private static List<string> BuildK1FormulaStrings(out int count)
    {
        const int dataRows = 100_000;
        const int formulaRows = 200_000;
        var list = new List<string>(formulaRows * 2);
        for (var r = 1; r <= formulaRows; r++)
        {
            var v = (r % dataRows) + 1;
            list.Add($"=IF(Data!A{v}=\"Show\",Data!B{v}*2,Data!B{v}/2)");
            list.Add($"=Data!B{v}*2+1");
        }

        count = list.Count;
        return list;
    }

    // The (id, sheetName) pairs the parser would allocate per reference token across the K1 formulas.
    private static List<(string Id, string SheetName)> CountReferenceStrings(out int refCount)
    {
        const int dataRows = 100_000;
        const int formulaRows = 200_000;
        var list = new List<(string, string)>(formulaRows * 4);
        for (var r = 1; r <= formulaRows; r++)
        {
            var v = (r % dataRows) + 1;
            // C{r}: A{v}, B{v}, B{v}
            list.Add(("A" + v, "Data"));
            list.Add(("B" + v, "Data"));
            list.Add(("B" + v, "Data"));
            // D{r}: B{v}
            list.Add(("B" + v, "Data"));
        }

        refCount = list.Count;
        return list;
    }

    // Sizes the Id and SheetName strings a CellReference tree holds. Also counts distinct sheet names (what
    // interning could collapse them to). Returns the reference-node count.
    private static long CollectCellReferenceStrings(
        Workbook wb,
        out long idBytes,
        out long sheetBytes,
        out int distinctSheetNames
    )
    {
        const int dataRows = 100_000;
        const int formulaRows = 200_000;
        long ids = 0;
        long sheets = 0;
        long refs = 0;
        var distinct = new HashSet<string>(StringComparer.Ordinal);
        for (var r = 1; r <= formulaRows; r++)
        {
            var v = (r % dataRows) + 1;
            var aLen = ("A" + v).Length;
            var bLen = ("B" + v).Length;
            // C{r}: A{v}, B{v}, B{v} ; D{r}: B{v}
            ids += ApproxStringBytes(aLen) + ApproxStringBytes(bLen) * 3;
            sheets += ApproxStringBytes(4) * 4; // "Data" per node
            refs += 4;
            distinct.Add("Data");
        }

        idBytes = ids;
        sheetBytes = sheets;
        distinctSheetNames = distinct.Count;
        return refs;
    }

    // ---------------------------------------------------------------------------------- index codec

    private static byte[] SerializeIndex(List<(int Col, int[] Rows)> columns)
    {
        var size = 4; // column count
        foreach (var (_, rows) in columns)
        {
            size += 4 + 4 + rows.Length * 4; // col + rowCount + rows
        }

        var buffer = new byte[size];
        var span = buffer.AsSpan();
        var offset = 0;
        BinaryPrimitives.WriteInt32LittleEndian(span.Slice(offset, 4), columns.Count);
        offset += 4;
        foreach (var (col, rows) in columns)
        {
            BinaryPrimitives.WriteInt32LittleEndian(span.Slice(offset, 4), col);
            offset += 4;
            BinaryPrimitives.WriteInt32LittleEndian(span.Slice(offset, 4), rows.Length);
            offset += 4;
            foreach (var row in rows)
            {
                BinaryPrimitives.WriteInt32LittleEndian(span.Slice(offset, 4), row);
                offset += 4;
            }
        }

        return buffer;
    }

    private static Dictionary<int, List<int>> DeserializeIndex(byte[] bytes, out int totalCells)
    {
        var span = bytes.AsSpan();
        var offset = 0;
        var colCount = BinaryPrimitives.ReadInt32LittleEndian(span.Slice(offset, 4));
        offset += 4;
        var map = new Dictionary<int, List<int>>(colCount);
        var total = 0;
        for (var c = 0; c < colCount; c++)
        {
            var col = BinaryPrimitives.ReadInt32LittleEndian(span.Slice(offset, 4));
            offset += 4;
            var rowCount = BinaryPrimitives.ReadInt32LittleEndian(span.Slice(offset, 4));
            offset += 4;
            var rows = new List<int>(rowCount);
            for (var i = 0; i < rowCount; i++)
            {
                rows.Add(BinaryPrimitives.ReadInt32LittleEndian(span.Slice(offset, 4)));
                offset += 4;
            }

            map[col] = rows;
            total += rowCount;
        }

        totalCells = total;
        return map;
    }

    private static byte[] BrotliBytes(byte[] raw)
    {
        using var output = new MemoryStream();
        using (var brotli = new BrotliStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            brotli.Write(raw);
        }

        return output.ToArray();
    }

    // ---------------------------------------------------------------------------------- helpers

    private static string ColumnName(int index)
    {
        var name = string.Empty;
        while (index > 0)
        {
            var remainder = (index - 1) % 26;
            name = (char)('A' + remainder) + name;
            index = (index - 1) / 26;
        }

        return name;
    }

    private static long ApproxStringBytes(int length)
    {
        var raw = 16 + 4 + 2L * length;
        return (raw + 7) & ~7L;
    }

    private static long Settle()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        return GC.GetTotalMemory(forceFullCollection: true);
    }

    private static void GcQuiesce()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}
