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
