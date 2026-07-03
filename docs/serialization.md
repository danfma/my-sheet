# Serialization (MemoryPack)

A `Workbook` serializes to a compact binary format via
[MemoryPack](https://github.com/Cysharp/MemoryPack). This is MySheet's *native* persistence — fast to
write, fast to load, and it round-trips the full expression trees (not just values). It is unrelated to
`.xlsx`; for Excel files see [Excel interop](excel-interop.md).

## Save and load

```csharp
using Danfma.MySheet;

workbook.Save("model.mysheet");
Workbook restored = Workbook.Load("model.mysheet");

// Async overloads:
await workbook.SaveAsync("model.mysheet", cancellationToken);
Workbook restoredAsync = await Workbook.LoadAsync("model.mysheet", cancellationToken);
```

`Load`/`LoadAsync` throw `InvalidDataException` if the file does not contain a workbook. The file
extension is yours to choose — the examples use `.mysheet` by convention.

## What round-trips — and what does not

The **cold** column is the default `Save`; the **warm** column is a save with
[`IncludeComputedValues`](#warm-start-persisting-computed-values) (everything a cold save persists,
plus the memoized values).

| | Cold | Warm | Notes |
| --- | --- | --- | --- |
| Sheets (name, tab order) | Yes | Yes | The case-insensitive name lookup is restored on deserialization. |
| Cells and full expression trees | Yes | Yes | Formulas stay formulas — a loaded workbook keeps recalculating. |
| Custom-function **calls** (`FunctionCall` nodes) | Yes | Yes | Name and argument expressions round-trip. |
| Custom-function **implementations** (delegates) | **No** | **No** | Behavior is code, not data — re-register after loading. |
| Memoization cache | **No** | **Partly** | Cold recomputes lazily on first read. Warm restores the cache — except volatile and reference-typed cells (below), which still recompute. |

The practical consequence: if your workbook uses [custom functions](custom-functions.md), re-register
them after every `Load`, or those calls evaluate to `#NAME?`:

```csharp
var restored = Workbook.Load("model.mysheet");

restored.RegisterFunction("CUSTOM", (arguments, wb) =>
{
    var a = arguments[0].Evaluate(wb).AsDouble() ?? 0;
    var b = arguments[1].Evaluate(wb).AsDouble() ?? 0;

    return a + b;
});

double value = restored.GetCellValue("Sheet1", "A1").ToDouble();
```

## Warm start: persisting computed values

By default a saved file is the **model only** — every value is recomputed lazily on the first read after
loading. Pass `WorkbookSaveOptions { IncludeComputedValues = true }` to also persist the memoization cache,
so a load starts **warm** and serves already-computed cells without re-evaluating them:

```csharp
workbook.Save("model.mysheet", new WorkbookSaveOptions { IncludeComputedValues = true });
// await workbook.SaveAsync("model.mysheet", new WorkbookSaveOptions { IncludeComputedValues = true }, ct);

var warm = Workbook.Load("model.mysheet"); // reads back with the cache pre-populated
```

`Load`/`LoadAsync` need no flag — they detect the format from the file header.

### File format

- **Cold** (`Save(path)`, or `IncludeComputedValues = false`) — the raw MemoryPack of the model, byte-for-byte
  identical to every prior version. This is a permanent contract, guarded by a regression test.
- **Warm** — a small self-describing container: the magic `MSWM`, a 1-byte format version, the **same** model
  bytes a cold save would write, then a value block (the MemoryPack of the cached values). `Load` sniffs the
  4-byte magic: a match is a warm container, anything else is a raw (cold or pre-existing) model, so old files
  keep loading unchanged.

Because the model and its values travel in one file, they can never desynchronize on load.

### What warm start does *not* freeze

Two kinds of cached value are deliberately **excluded** from the snapshot and recompute on first read, even
from a warm file:

- **Volatile cells** — anything that touched `NOW`/`TODAY`/`RAND`/`RANDBETWEEN` (directly or transitively).
  Persisting them would "freeze yesterday's clock"; instead they re-sample on the next read.
- **Reference-typed results** — rare as a final cell value and cheap to rebuild.

### Staleness contract

Warm start persists values you already computed; it does not track edits. The post-load contract is the same
as always: **after editing cells, call `InvalidateCache()`** (or `Recalculate()` for a volatile-only refresh)
before reading, or you will read stale values. A warm load only skips the *first* recomputation of unchanged,
non-volatile cells — it changes nothing about how invalidation works afterwards. And, as with a cold load,
[custom functions](custom-functions.md) must still be re-registered: cells that were **not** cached at save
time (or that you invalidate) will re-evaluate their calls and need the implementation present.

## Compatibility

Expression nodes are serialized as a MemoryPack union, and the union tags are **append-only by
project policy**: existing tags are never renumbered, reordered or reused, and new node types get new
tags. Workbooks saved by an older version therefore remain loadable by newer versions of the library.

Because only the tags (never type names) go on the wire, the [2.0 namespace
reorganization](migrating-to-2.0.md) did not change the format at all: files saved by 1.x load in 2.0
unchanged, guarded by a frozen pre-2.0 binary fixture in the test suite.

## When to use which format

| Need | Use |
| --- | --- |
| Fast native persistence of a computed model (cache-style, service restarts, snapshots between processing steps) | `Workbook.Save` / `Load` |
| Interchange with people or other tools (open in Excel, send a report) | [`SaveAsExcel` / `MergeIntoExcel`](excel-interop.md) |
| Ingesting the source-of-truth spreadsheet | [`ExcelFile.Load`](excel-interop.md) |
