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
- The memoization cache stores the `ComputedValue` struct **inline** in the dictionary — no long-lived
  per-cell box, which is what previously generated Gen1 GC pressure.
- Range-consuming functions (`SUM(A1:A1000)`, lookups, …) enumerate cell values through the cache with
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

`Workbook.GetCellValue(sheetName, id)` caches each cell's `ComputedValue` under `(sheet, id)`. Every
`CellReference` inside a formula resolves through the same cache, so a cell referenced by many formulas
— or reached both directly and through a range expansion — is **computed exactly once**:

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

The cache is a `ConcurrentDictionary`, safe for concurrent background readers. It is not serialized —
a loaded workbook starts cold and fills lazily.

### Circular references

The memoization layer tracks the cells being evaluated on the current thread. A cycle (`A1=B1`,
`B1=A1`) is detected and returns `#REF!` (`Error.Ref`) instead of overflowing the stack. The tracking
is thread-local, so concurrent evaluation of the same cell on different threads is not a false cycle.

## Whole-column references at scale

Formulas that consume a **whole-column reference** — `MATCH(x, A:A)`, `VLOOKUP(x, A:B, 2, FALSE)`,
`SUMIF(A:A, k)`, `SMALL(A:A, k)`, `SUM(A:A)` — are the pathological case for a naive engine. Each such
formula scans every populated cell of the column, so a sheet with *F* whole-column formulas over a
column of *N* cells costs **O(F·N)**. On a real workbook (~506k cells in one column, ~400k formulas
referencing it) that is ~2×10¹¹ visits — roughly **57 minutes** of pure scanning on a load-once /
read-once pass. Two internal caches collapse that to seconds. Both are bounded, discardable, and add
**no public API** — the sparse model and the single-cell hot path are untouched.

### Layer 1 — the structural index

Per sheet, a lazy index maps `column → cell ids ordered by row` (and the symmetric `row → ids` on
demand). It answers "which cells does `A:A` cover, in order?" without re-scanning the cell dictionary.
Range bound resolution (the table lookup `VLOOKUP`/`INDEX`/`OFFSET` do on every evaluation) reads the
bounding box straight from the ordered lists by binary search.

This is what makes **small columns in a big sheet** cheap once the index exists: before the index, a
formula referencing a 16-cell column still walked all ~1.5M keys of the sheet to find those 16; after
it, only the 16 are visited.

**Second-use admission.** Building the index is itself an O(N) pass (plus a bucket per populated column
or row). That pays off only when a sheet is read through open ranges **more than once in an epoch**. A
sheet read *exactly once* — the "`InvalidateCache()` then one whole-column read per epoch" shape — would
pay a whole-sheet index build just to serve one column it never reuses. So, exactly like the range
snapshot below, the index is **admitted on a sheet's second open-range read, not its first**: the first
read serves itself from a **direct key scan** (the pre-index path), collecting and sorting *only the
matched ids* into the same deterministic order the index yields; the second read builds the index and
every read after reuses it. In a reuse-heavy pass (thousands of formulas over `A:A`) the single extra
scan on the first formula is invisible; in a single-read-per-epoch pass it stays at pre-index (2.6.1)
parity instead of rebuilding the index every epoch.

**Lazy per-bucket sort.** Building the index bucketizes the ids but does **not** sort them; each
column's list is ordered (by row) only on its first access, so a read that touches one narrow column of
a wide sheet sorts only that one list, not every column.

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

Both caches are per-workbook, `ConcurrentDictionary`-based, created race-free on first use, and **not
serialized** (a loaded workbook starts cold). They differ only in invalidation, because the structural
index describes *structure* and the range snapshot carries *values*:

| Cache | Built | `Recalculate()` | `InvalidateCache()` |
| --- | --- | --- | --- |
| Structural index (Layer 1) | **second** open-range read of a sheet | **survives** (structure ≠ values) | dropped |
| Range snapshots (Layer 2) | **second** cacheable range read | **dropped** (values may be volatile-tainted) | dropped |

```csharp
// Load once, then read a lot: the caches fill lazily on the first pass and every later read is served
// from them.
var total = workbook.GetCellValue("Calc", "A1");   // 1st read of A:A → marks it, linear path
var next  = workbook.GetCellValue("Calc", "A2");   // 2nd read → builds the shared snapshot
var more  = workbook.GetCellValue("Calc", "A3");   // 3rd+ read → served from the shared snapshot

// After editing cells, flush everything before reading again:
sheet["A1"] = new NumberValue(20);
workbook.InvalidateCache();                          // drops both layers

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

1. Read cells through `GetCellValue` (memoized), not by re-evaluating expressions in a loop.
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
