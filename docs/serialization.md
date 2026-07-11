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

## Save options

The `Save(path, WorkbookSaveOptions)` / `SaveAsync` overloads take two **orthogonal** switches. `Load` needs
no matching flag — it detects the format (raw vs. container, uncompressed vs. Brotli) from the file header.

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| [`IncludeComputedValues`](#warm-start-persisting-computed-values) | `bool` | `false` | Persist the memoization cache alongside the model so a load starts **warm** (skips recomputation). |
| [`Compression`](#compression) | `WorkbookCompression` | `None` | `Brotli` shrinks the file with BCL Brotli. |
| [`CompressionLevel`](#compression) | `CompressionLevel` | `Optimal` | Brotli quality when compressing. `Fastest` cuts save time markedly on large workbooks for a larger file; a write-time knob only — `Load` reads any level. |

With the first two at their defaults, `Save(path, options)` is byte-identical to `Save(path)`.

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

- **Cold, uncompressed** (`Save(path)`, or `IncludeComputedValues = false` with `Compression = None`) — the
  raw MemoryPack of the model, byte-for-byte identical to every prior version. This is a permanent contract,
  guarded by a regression test.
- **Container** — every other combination is a small self-describing container: the magic `MSWM`, a 1-byte
  format version, the uncompressed model length (int32 LE), then the body. `Load` sniffs the 4-byte magic: a
  match is a container, anything else is a raw (cold or pre-existing) model, so old files keep loading
  unchanged. The version byte selects the body encoding:
  - **v1 (uncompressed warm)** — the **same** model bytes a cold save would write, then a value block (the
    MemoryPack of the cached values). Warm-start files written before compression existed are exactly this.
  - **v2 (Brotli)** — the model and value block concatenated and Brotli-compressed as a *single* stream (one
    stream compresses better than two independent blocks). Used for any compressed save, cold or warm; a cold
    compressed file simply carries an empty value block.

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

## Compression

MemoryPack optimizes for speed, so its layout is fixed-width and redundant — which means it compresses
extremely well. Pass `WorkbookCompression.Brotli` to shrink the saved file with the BCL's Brotli
(`CompressionLevel.Optimal`); no third-party dependency is added.

```csharp
workbook.Save("model.mysheet.br", new WorkbookSaveOptions { Compression = WorkbookCompression.Brotli });

var restored = Workbook.Load("model.mysheet.br"); // detects and decompresses transparently
```

Compression is orthogonal to warm start — combine them to persist a warm cache in a compressed file:

```csharp
workbook.Save("model.mysheet.br", new WorkbookSaveOptions
{
    IncludeComputedValues = true,
    Compression = WorkbookCompression.Brotli,
});
```

### Measured sizes

Brotli-`Optimal` over the production MemoryPack bytes, three representative workbooks (Apple M1 Pro,
.NET 10). Percentages are the compressed size as a fraction of the raw MemoryPack file:

| Workbook | Cells | Raw MemoryPack | Brotli | Fraction |
| --- | ---: | ---: | ---: | ---: |
| Small (fixture-like) | 20 | 1,147 B | 289 B | ~25% |
| Medium (values + formulas) | 7,500 | 348,035 B | 33,626 B | ~10% |
| Large (whole-column model) | 302,048 | 7,935,568 B | 1,090,808 B | ~14% |

The larger and more repetitive the model, the bigger the win — a real workbook typically drops to well
under half its raw size. Compression trades CPU at save/load time for that space; leave it `None` when you
save frequently to a fast local disk and file size is not a concern.

### File naming

The library **never** renames the file you pass — a compressed save writes exactly the path you give it,
with no extension appended. Because the `MSWM` container is self-describing, `Load` does not rely on the
name to decide whether to decompress. If you want compressed files to be recognizable, adopt a suffix
convention in your own code (a `.br` suffix, as in the examples above, is the common choice).

## Compatibility

Expression nodes are serialized as a MemoryPack union, and the union tags are **append-only by
project policy**: existing tags are never renumbered, reordered or reused, and new node types get new
tags. Workbooks saved by an older version therefore remain loadable by newer versions of the library.

Because only the tags (never type names) go on the wire, the [2.0 namespace
reorganization](migrating-to-2.0.md) did not change the format at all: files saved by 1.x load in 2.0
unchanged, guarded by a frozen pre-2.0 binary fixture in the test suite.

### Forward-compatibility: shared-formula delta nodes (tags 319-321)

Shared formulas (dragged Excel formulas) can now be represented by three additional node types —
`AnchoredCellReference` (319), `AnchoredRangeReference` (320) and `SharedFormulaSlave` (321) — that let
every slave cell of a supported group share one master expression tree instead of holding an independent,
fully-expanded one (see [Excel interop → Shared
formulas](excel-interop.md#shared-formulas-a-shared-master-tree-with-per-slave-deltas) for what makes a
group "supported" and the measured load-time win).

This is a **one-way** compatibility boundary, same as any append-only tag addition:

- A file saved by **this or a later** version of the library — whether produced by `Workbook.Save` or by
  `ExcelFile.Load` followed by a save — can contain cells using tags 319-321 whenever the workbook holds a
  supported shared-formula group. Such a file **cannot be opened by a version of the library older than
  the one that introduced these tags**: the older MemoryPack union does not recognize them and deserialization
  fails.
- A file saved by an **older** version of the library never contains these tags, and continues to load
  unchanged in this and every later version, exactly as the append-only policy above guarantees.

**Honest note: this is a RAM/GC optimization, not a disk-size one.** In memory, every slave in a supported
group shares a single `Expression` instance for its master tree — that is where the allocation and GC win
comes from. On the wire, MemoryPack serializes each node's data independently and does **not** perform
reference-tracking or structural deduplication: a `SharedFormulaSlave` still writes its own copy of the
master tree's serialized bytes, once per slave. A workbook with a large shared-formula group therefore does
not shrink on disk from this change alone — only its in-memory footprint after loading does.

## When to use which format

| Need | Use |
| --- | --- |
| Fast native persistence of a computed model (cache-style, service restarts, snapshots between processing steps) | `Workbook.Save` / `Load` |
| Interchange with people or other tools (open in Excel, send a report) | [`SaveAsExcel` / `MergeIntoExcel`](excel-interop.md) |
| Ingesting the source-of-truth spreadsheet | [`ExcelFile.Load`](excel-interop.md) |
