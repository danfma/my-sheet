# Performance

MySheet's design goal is a formula engine cheap enough to run on a server hot path: load a workbook
once, then evaluate and re-evaluate many cells with minimal allocation and no per-read setup cost. This
guide explains the three mechanisms that make that work — the allocation-free `ComputedValue` union,
per-cell memoization, and `RunWithLargeStack` — and reports the numbers actually measured in this
repository.

## Allocation-free evaluation

`Expression.Evaluate` returns a [`ComputedValue`](computed-value.md): a `readonly struct` union of one
`double`, one `object?` and a one-byte tag. Numbers, booleans, blanks and errors are carried entirely
inside the struct — **no boxing, no heap allocation**. Text and references only carry a pointer to a
string or node that already existed.

Supporting choices in the same spirit:

- `EvaluationContext` is a `readonly struct`, so threading the context through the recursive evaluation
  allocates nothing (only `LET` name bindings, when used, live on the heap).
- The memoization store keeps the `ComputedValue` struct **inline** in a dense per-sheet page (see
  [the value store](#the-value-store-a-dense-paged-cache) below) — no long-lived per-cell box, no
  per-entry hash node, which is what previously generated Gen1 GC pressure.
- Range-consuming functions (`SUM(A1:A1000)`, lookups, …) enumerate cell values through the store with
  the same non-boxing type.

### Measured numbers

The `ComputedValue` design was adopted based on a BenchmarkDotNet experiment in this repository
(`plans/cellvalue-boxing-experiment/`, runnable spike in `benchmarks/Danfma.MySheet.Benchmark`)
comparing the previous `object?`-returning evaluator against the struct union. Environment: Apple M1
Pro, .NET 10, BenchmarkDotNet ShortRun. Highlights (ratios vs. the boxed `object?` baseline):

| Workload | Boxed `object?` | `ComputedValue` | Allocation |
| --- | --- | --- | --- |
| Cumulative dependency chain, 3,000 cells | 44.4 μs, 144 KB | 18.4 μs (**0.42×** time) | **0 B** (was 144 KB) |
| `SUM` fold over 100,000 values | 476.9 μs, 2.4 MB | 111.8 μs (**0.23×** time) | **0 B** (was 2.4 MB) |
| Mixed arithmetic, 100,000 ops | 611.7 μs, 2.28 MB | 202.4 μs (**0.33×** time) | **0 B** (was 2.28 MB) |
| Cache-heavy graph, 100,000 cells | 5.64 ms, 2.4 MB | 4.95 ms (**0.88×** time) | **0 B** (was 2.4 MB) |

In short: **2.3–4.3× faster on arithmetic-heavy paths, 4–12% faster on cache-heavy paths, and zero
GC collections in all cases**. Two honesty notes: these compare MySheet's own design alternatives, not
other libraries; and the raw tables (including 1,000-cell runs) are in
`plans/cellvalue-boxing-experiment/results-*.md` if you want the full data. The elimination of
allocations — zero Gen0/Gen1 activity in a background extractor — was the decisive argument, more than
raw throughput.

## Memoization

`Workbook.GetCellValue(sheetName, id)` memoizes each cell's `ComputedValue` in the workbook's value
store, addressed by a sheet handle plus the cell's `(column, row)`. Every `CellReference` inside a
formula resolves through the same store, so a cell referenced by many formulas — or reached both
directly and through a range expansion — is **computed exactly once**:

```csharp
sheet["A1"] = ExpressionParser.Parse("=EXPENSIVE()", sheet);
sheet["B1"] = ExpressionParser.Parse("=A1+A1", sheet);
sheet["C1"] = ExpressionParser.Parse("=SUM(A1:A1)+A1", sheet);

workbook.GetCellValue("Sheet1", "B1");   // EXPENSIVE() runs once
workbook.GetCellValue("Sheet1", "C1");   // ...and is not run again here
```

### Invalidation is explicit

The cache is **never invalidated automatically**. After mutating cells, call `InvalidateCache()`
(a full flush) before reading again:

```csharp
sheet["A1"] = new NumberValue(20);
// Without this, reads still serve the previously cached values:
workbook.InvalidateCache();
```

Why this design (decided and documented in `plans/memoization.md`):

- The target workload is **read-heavy extraction with rare, batched mutations** — edit, invalidate
  once, then read a lot.
- Surgical dependency-graph invalidation was evaluated and rejected: ranges make the reverse graph
  expensive (`SUM(A1:A1000)` alone is 1,000 edges) for little gain under rare mutations.
- The known trade-off: each reference adds one dictionary lookup, a small cost on trivial cells that is
  repaid whenever cells are shared or expensive.

### The value store: a dense paged cache

The memoized values do not live in a hash map keyed by `(sheet, id)` strings. They live in a **dense,
paged store addressed numerically** — the shape the compute-phase profiling pointed at. A string-keyed
`ConcurrentDictionary` paid, on every miss, a bucket-lock insert, a string-tuple hash, and one heap node
per entry (Gen1 pressure); the dense store removes all three from the hot path.

**Numeric addressing, derived on the fly.** `GetCellValue` resolves the sheet name to a small integer
handle once, then derives the cell's `(column, row)` straight from its canonical A1 id without allocating
(`CellAddress.TryGetColumnRow` — no substring, no `int.Parse`). The public API and the on-disk format are
untouched: cells are still addressed by A1 string on the surface, and the store is runtime-only, never
serialized (a loaded workbook starts cold and fills lazily).

**Two-level paging on both axes.** Per sheet, a slab holds a two-level column directory —
`groups[col >> 6][col & 63]` (groups of 64 columns) → `column.pages[row >> 10]` → a page of 1,024
`ComputedValue` slots plus a small presence bitmap. A lone high-index gridless column (e.g. `AAAA`) then
costs one group, not a giant flat array, and the hot read is pure shift/mask plus a couple of array
dereferences — no hashing. Directories start tiny and grow by doubling, publishing each new array with a
`Volatile` write. The presence bitmap is explicit because a zeroed slot is ambiguous — "not computed
yet" versus "computed to blank".

**Adaptive first page.** A full-size page allocated for a column that only ever holds a handful of cells
is mostly waste. So a column's **first** page is born with a small slot array covering the same logical
row span, and is promoted — reallocated toward the full page size — only when a write lands beyond its
current length; the promotion happens inside the page's seqlock, so a reader either sees the old array or
retries. A column that overflows its first page has proven it is dense, so its **later** pages are born
full — the small birth trims tiny and sparse sheets without adding reallocation churn to the dense path.
The initial size is the `InitialPageSlots` knob on `ValueStoreOptions` (a power of two, defaulting small).

**Lock-free reads via a per-page seqlock.** A `ComputedValue` is a multi-word struct, so a naive
concurrent read could tear. Each page is a **seqlock**: readers are lock-free and simply re-read if a
writer touched the page mid-read (an odd version number), while a single writer per page serialises on a
CAS gate and bumps the version around its update. Because memoized cells are read far more than written,
this keeps the read path free of the per-read monitor cost the old dictionary charged — evaluation stays
safe to run concurrently (as the volatile/F1 paths require).

**Sparsity guard.** A page is a fixed block, so cells scattered one-per-page over a huge row span would
balloon the footprint. Each slab watches its pages-allocated versus cells-present ratio; once a slab is
provably too sparse, it stops allocating **new** pages and diverts further scattered cells to a per-slab
dictionary. Dense sheets never trip it; a pathological scatter degrades to the dictionary's footprint
instead of the balloon. Nothing migrates and no page is torn down, so the concurrent read path is
untouched.

**Epoch semantics are preserved exactly.** `InvalidateCache()` drops every memoized value (the pages are
cleared, the store and its handle map kept); `Recalculate()` drops only the volatile-tainted cells so
they recompute lazily on the next read. The tainted set is a sparse keyed collection (only volatile cells
land in it), and the value epoch is independent of the write-maintained structural index below, which
survives both calls.

**Tunable geometry (defaults ship tuned for the K1 workload).** The row-page size, the column-group size,
and the two sparsity-guard thresholds are configurable per workbook through `WorkbookOptions.ValueStore`.
The defaults (row page 1,024, column group 64, warm-up 64 pages, floor 4 cells/page) reproduce the shipped
behavior, so `new Workbook()` is unchanged. A smaller page wastes less on tiny or sparse sheets; a larger
one scans faster on dense ones. Both sizes must be powers of two — the hot path is shift/mask, and the
shift is derived from the size — so a non-power-of-two or out-of-range value throws `ArgumentException` at
construction. The options are runtime **configuration**, not document state: they are never serialized (the
file format is untouched), and a workbook loaded from disk falls back to the defaults.

```csharp
// A workbook of many small sheets: shrink the page so each one wastes less memory.
var workbook = new Workbook(new WorkbookOptions
{
    ValueStore = new ValueStoreOptions
    {
        RowPageSize = 256,       // power of two in [64, 65536]  (default 1024)
        ColumnGroupSize = 32,    // power of two in [8, 4096]    (default 64)
        SparsityWarmupPages = 64,     // pages before density is judged  (default 64)
        SparsityMinCellsPerPage = 4,  // sparsity floor, in [1, RowPageSize]  (default 4)
    },
});
```

### String hygiene in the model

Two internal refinements keep the resident model lean without touching the file format. Sheet
qualifiers in references (`Data!B2`) are interned — every reference to the same sheet shares one
string instance, both when parsed and when loaded from a file. And the cell store keys its canonical
A1 cells by numeric `(column, row)` internally, keeping non-canonical ids (anything that does not
round-trip through the A1 writer) in a string overflow so the public string surface and the
serialized bytes stay exactly as before.

### Circular references

The memoization layer tracks the cells being evaluated on the current thread. A cycle (`A1=B1`,
`B1=A1`) is detected and returns `#REF!` (`Error.Ref`) instead of overflowing the stack. The tracking
is thread-local, so concurrent evaluation of the same cell on different threads is not a false cycle.

## Bulk extraction: `GetValueReader`

`GetCellValue(sheetName, id)` is the right default read, but an extraction loop that builds ids
(`"C" + row`) pays an id-string allocation, an A1 parse and a sheet-name hash **per cell**.
[`Workbook.GetValueReader`](workbook-and-expressions.md#bulk-reads-getvaluereader) resolves the sheet
handle once and reads by numeric address with identical semantics (misses evaluate on demand, literals
included):

```csharp
var reader = workbook.GetValueReader("Results");
var value = reader.GetValue(column: 3, row: 42); // no strings on a cache hit
```

Measured (360k computed cells): `29.8 ms / 24.2 MB` allocated with per-cell id strings →
`6.9 ms / 0 bytes` with the reader — 4.3x faster, allocation-free.

## Whole-column references at scale

Formulas that consume a **whole-column reference** — `MATCH(x, A:A)`, `VLOOKUP(x, A:B, 2, FALSE)`,
`SUMIF(A:A, k)`, `SMALL(A:A, k)`, `SUM(A:A)` — are the pathological case for a naive engine. Each such
formula scans every populated cell of the column, so a sheet with *F* whole-column formulas over a
column of *N* cells costs **O(F·N)**. On a real workbook (~506k cells in one column, ~400k formulas
referencing it) that is ~2×10¹¹ visits — roughly **57 minutes** of pure scanning on a load-once /
read-once pass. Two internal caches collapse that to seconds. Both are bounded and internal — no cache
tuning knobs to set — and the sparse model and the single-cell hot path are untouched.

### Layer 1 — the structural index

Per sheet, the index maps `column → the rows populated in it, ordered` (and the symmetric
`row → columns` on demand). It answers "which cells does `A:A` cover, in order?" without re-scanning the
cell dictionary. Range bound resolution (the table lookup `VLOOKUP`/`INDEX`/`OFFSET` do on every
evaluation) reads the bounding box straight from the ordered lists by binary search.

The buckets hold the coordinate as a plain integer — a column keeps its row numbers, a row its column
numbers — never the A1 id string. So the open-range value path expands straight from the stored
`(column, row)` into the dense store with no id to parse, and the index itself carries only integers: it
is markedly smaller in memory than an id-per-cell form, and far more compact still if it is ever
persisted. An id is re-derived only on the cold enumeration that yields cell *expressions* (not values).

This is what makes **small columns in a big sheet** cheap: before the index, a formula referencing a
16-cell column still walked all ~1.5M keys of the sheet to find those 16; after it, only the 16 are
visited.

**Write-maintained and lifetime-scoped.** The index is built **once per sheet, lazily on that sheet's
first open-range read**, and then kept current by every cell write. The `Sheet` funnels all mutation
through a single write choke point — the indexer `set` and `Remove` — and each insert or delete updates
the index in place, so it never needs a full rebuild. Because it is maintained on the write, it is *not*
tied to a value-cache epoch: unlike the value caches, **`InvalidateCache()` does not drop it** (clearing
memoized values never changes which cells exist — only a write does, and a write maintains it). Absorbing
a write is cheap: an in-order append — the common Fill pattern — is O(1); an out-of-order insert only
marks its one column dirty so that column re-sorts on its next read; a delete removes the id from its
bucket in O(bucket). The index is never serialized, so a workbook loaded from disk starts without one and
rebuilds it once on its first open-range read for that instance's life.

**Lazy per-bucket sort.** Building the index bucketizes the coordinates but does **not** sort them; each
column's list is ordered (by row) only on its first access, so a read that touches one narrow column of
a wide sheet sorts only that one list, not every column. The ordering is numeric by that secondary axis
(a column by row number, a row by column number) — `A2` precedes `A10`, not the reverse a text sort of
the ids would give.

### Layer 2 — epoch range caches

For a populated range above a size threshold that is read **more than once in an epoch**, the engine
materializes a **snapshot** — the range's `ComputedValue`s in enumeration order, read exactly once
through Layer 1 + the memoized cell cache. Every accelerator is then built **lazily, on first demand,
and shared by every formula of the epoch**:

- an **exact-value hash** (`value → first position`) → `MATCH(…,0)`, exact `XLOOKUP`/`XMATCH`,
  `VLOOKUP(…,FALSE)` in O(1);
- a **sorted index** with prefix/suffix-max positions → approximate `MATCH` type 1/-1 and
  `VLOOKUP(…,TRUE)` by binary search, reproducing Excel's "last of the tied" tie-break for any input
  order;
- a **sorted-number view** → `SMALL`/`LARGE`/`MEDIAN`/`PERCENTILE`/`QUARTILE` by direct indexing;
- a **numeric-equality map** (`value → (sum, count)`) → equality `SUMIF`/`COUNTIF`/`AVERAGEIF` in O(1);
- an **aggregate memo** → `SUM`/`COUNT`/`MAX`/`MIN`/`AVERAGE` of a repeated pure range folded once.

The overall cost drops from O(F·N) to **O(N + F·log N)**. Semantics are preserved bit for bit: where a
value does not fit an index cleanly (blank-equivalent lookups, wildcard/comparison criteria, the
"closest" approximate `XLOOKUP`) the consumer falls back to a **linear scan over the cached snapshot** —
still cache-served, so the O(N) re-read of cell values is gone even on the fallback path.

**How the snapshot is materialized.** Because a closed range knows its bounds up front, the snapshot is
written into a result array sized exactly once — no growing intermediate list, no final copy. When every
page the range covers is fully present, the values are lifted a page at a time with a block copy straight
out of the store rather than cell by cell, with the page's seqlock version re-checked after each copy so a
concurrent write can never smuggle a torn value into the snapshot (a page that fails the re-check falls
back to the on-demand cell path). The allocation then amounts to the result array alone.

**Second-use admission.** Materializing a snapshot only pays off when a range is read repeatedly.
A range read *exactly once* per epoch — a sliding window (`SUM(A$1:A500)`, `SUM(A$1:A501)`, …), a
one-shot bounded lookup, an invalidate-heavy loop — would pay an O(N) build it never reuses. So a
range is **admitted on its second read, not its first**: the first read only records a lightweight
marker and takes the linear path; the snapshot is built on the second read and shared by every read
thereafter. In a reuse-heavy pass (thousands of formulas over `A:A`) this costs one extra linear scan
on the very first formula and is invisible; in a single-use pass it removes the wasted materialization
entirely. (A marker is dropped with the cache at epoch end; a defensive cap of 64k markers stops an
adversarial flood of distinct single-use ranges from growing the marker set without bound.)

**The 256-cell threshold.** A range with fewer than 256 populated cells is not cached — and not even
marked: a linear scan already wins there and tracking every tiny range would only flood the dictionary.
The check uses a cheap upper-bound estimate (rectangle area, or the sum of the covered structural-index
lists) — it never materializes a snapshot just to decide.

### Lifecycle: when the caches are built and dropped

Both caches are created race-free on first use and **not serialized** (a loaded workbook starts cold).
They differ in ownership and invalidation, because the structural index describes *structure* (owned by
the `Sheet`, maintained on the write) while the range snapshot carries *values* (owned by the `Workbook`,
epoch-scoped):

| Cache | Built | `Recalculate()` | `InvalidateCache()` |
| --- | --- | --- | --- |
| Structural index (Layer 1) | **first** open-range read of a sheet (once per its life), then write-maintained | **survives** (structure ≠ values) | **survives** (write-maintained, not epoch-scoped) |
| Range snapshots (Layer 2) | **second** cacheable range read | **dropped** (values may be volatile-tainted) | dropped |

```csharp
// Load once, then read a lot: the caches fill lazily on the first pass and every later read is served
// from them.
var total = workbook.GetCellValue("Calc", "A1");   // 1st read of A:A → marks it, linear path
var next  = workbook.GetCellValue("Calc", "A2");   // 2nd read → builds the shared snapshot
var more  = workbook.GetCellValue("Calc", "A3");   // 3rd+ read → served from the shared snapshot

// After editing cells, flush the value caches before reading again. The write already maintained the
// structural index, so InvalidateCache leaves Layer 1 in place and drops only the value caches:
sheet["A1"] = new NumberValue(20);
workbook.InvalidateCache();                          // drops the value caches; Layer 1 survives

// A volatile refresh (NOW/RAND) drops only the value snapshots; the structural index survives:
workbook.Recalculate();
```

### Measured numbers

Synthetic whole-column benchmark (`benchmarks/Danfma.MySheet.Benchmark`, `--whole-column-scale`), Apple
Silicon, .NET 10. Each row is one `InvalidateCache()` + one full evaluation pass (the load-once /
read-once cycle). "Before" is the pre-cache engine; "After" is the two-layer engine.

**Reduced (50k data cells × 10k formulas), formula block over the big column, wall-clock:**

| Formula | Before | After |
| --- | --- | --- |
| `MATCH(…,1)` | 80.5 s | 141 ms |
| `MATCH(…,0)` | 73.6 s | 95 ms |
| `VLOOKUP(…,FALSE)` | 8.1 s | 155 ms |
| `SUMIF` (equality) | 70.5 s | 68 ms |
| `SMALL` | 59.5 s | 75 ms |
| `SUM` (repeated) | 55.2 s | 89 ms |

**Full (500k data cells ≈ 1.5M cells in the sheet × 100k formulas), measured, not extrapolated:** any
single block of 100k whole-column formulas over the big column runs in **≤1.15 s** (down from an
estimated **~4.2 hours** on the pre-cache engine — ~18,000× on the pure-scan functions). All 14 blocks
(7 functions × 2 targets, 1.4M evaluations) total **7.7 s**. The extra memory the evaluation retains
(cell memo + structural index + range snapshot) is ~75–122 MB for a 500k-cell big column, transient and
dropped on invalidation; the process peak is dominated by the workbook itself, not the caches.

> **Measure, don't extrapolate.** The pre-cache cost was truly O(F·N), so timing 1k formulas and
> multiplying by 100 was a valid estimate. With the caches the cost is O(N + F·log N): the snapshot
> build is a one-time O(N) cost amortized across the whole block, so a 1k sample × 100 multiplies that
> one-time build by 100 and wildly over-estimates. The full number above is measured over the real 100k
> formulas.

### Repeated whole-column reads scale with the column, not the sheet

Because Layer 1 is write-maintained and survives `InvalidateCache()`, the "load once, then re-read a
whole column every epoch" shape costs the same no matter how large the *rest* of the sheet is. This
harness (`--structural-index-lifetime`) fixes column A at 200 populated cells, grows the sheet total from
10k to 500k by filling *other* columns, and times `{ InvalidateCache(); COUNTIF(A:A,">0") }` per epoch:

| Sheet total cells | Mean ms per epoch | First read (one-time index build) |
| --- | --- | --- |
| 10,200 | 0.31 | 12.2 ms |
| 40,200 | 0.29 | 1.4 ms |
| 100,200 | 0.28 | 5.2 ms |
| 200,200 | 0.11 | 14.6 ms |
| 500,200 | 0.11 | 30.4 ms |

The per-epoch time stays flat (it tracks column A's 200 cells, not the sheet), while the *first* read
pays the one-time O(sheet) index build. The index survives every later `InvalidateCache()`, so no epoch
after the first rebuilds it — the read is served from the maintained index each time. For reference, the
comparison the design targeted (Aspose's write-time column model) is ~0.4 ms constant on this shape.

### Compared to ClosedXML

The same reduced whole-column workload (50k × 10k), each engine in its own process:

| Formula | MySheet | ClosedXML |
| --- | --- | --- |
| `MATCH(…,1)` (approximate) | 123 ms | 67,294 ms |
| `MATCH(…,0)` (exact) | 105 ms | 1,821 ms |
| `VLOOKUP(…,FALSE)` | 154 ms | 1,463 ms |
| `SUMIF` (equality) | 95 ms | 16,874 ms |
| `COUNTIF` (equality) | 97 ms | 16,970 ms |
| `SUM` (repeated) | 113 ms | 20,192 ms |
| `SMALL` | ~78 ms | **`#NAME?` — not implemented** |

Results are identical where both engines compute. MySheet is faster in every supported function (9.5× to
547×) **and** holds a lower peak working set (~152 MB vs ~156 MB) even while carrying the epoch caches —
and it answers `SMALL`/`LARGE` over a whole column, which ClosedXML does not evaluate at all.

## Deep chains: `RunWithLargeStack`

Evaluation is recursive, and the risk is not deep *formulas* but long **dependency chains between
cells** — e.g. a cumulative column (`B2=B1+A2`, `B3=B2+A3`, …) thousands of rows deep. Computing the
last cell recurses through the whole chain, and memoization does not help the *first* computation. On a
default ~1 MB thread stack this overflows after a few thousand frames — and .NET cannot catch a
`StackOverflowException`.

The solution is to run the evaluation batch on a thread with a large reserved stack:

```csharp
var value = Workbook.RunWithLargeStack(() => workbook.GetCellValue("Sheet1", "A20000"));

// Or wrap a whole extraction batch — the thread cost is paid once, not per cell:
var totals = Workbook.RunWithLargeStack(() =>
{
    var results = new Dictionary<string, ComputedValue>();

    foreach (var id in workbook["Sheet1"].Keys)
    {
        results[id] = workbook.GetCellValue("Sheet1", id);
    }

    return results;
});
```

- The default stack size is 256 MB, and it is a **reservation**: physical memory grows only with the
  depth actually reached. An optional second parameter overrides the size.
- Exceptions thrown inside the work are captured and re-thrown on the caller with their original stack
  trace.
- A 20,000-cell cumulative chain — which overflows the default stack — is covered by the test suite
  through this API.
- The Excel exporter and merger do this for you: `SaveAsExcel` and `MergeIntoExcel` evaluate all cells
  up front inside one `RunWithLargeStack` call.

## Practical checklist

1. Read cells through `GetCellValue` (memoized), not by re-evaluating expressions in a loop; for
   large extraction loops, switch to `GetValueReader` (numeric addresses — no per-cell strings).
2. Batch your mutations, then call `InvalidateCache()` once.
3. Wrap large extraction batches in a single `Workbook.RunWithLargeStack(...)` call.
4. Extract results with `TryGet*`/`To*`; keep `AsObject()` (which boxes) out of hot loops.
5. Custom functions are cached per cell like built-ins — don't rely on re-execution per read
   ([details](custom-functions.md#interaction-with-memoization)).

## Benchmarks in the repository

`benchmarks/Danfma.MySheet.Benchmark` is a BenchmarkDotNet project with the engine benchmarks (parsing,
in-memory workbook construction and evaluation, with ClosedXML as an independent in-memory reference
point) and the `ComputedValue` spike:

```shell
dotnet run --project benchmarks/Danfma.MySheet.Benchmark -c Release
```

The whole-column scale numbers above come from a dedicated wall-clock harness (the O(F·N) baseline is
too slow to iterate under BenchmarkDotNet):

```shell
# Reduced scale (50k × 10k), runs in seconds:
dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale
# + the measured full scale (500k × 100k), with per-block memory:
dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --whole-column-scale --full
```
