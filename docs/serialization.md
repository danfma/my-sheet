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
| Memoization cache | **No** | **Partly** | Cold recomputes lazily on first read. Warm restores the cache — except volatile cells (below), which still recompute. |

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
  raw MemoryPack of the model, with **no container header of any kind**. That *write shape* is the permanent
  contract, and it is deterministic: the same model always serializes to the same bytes, guarded by frozen
  base64 goldens in the test suite (`CellStoreTests.Wire_IsByteIdentical_AfterNumericKeys` and
  `SheetNameInterningTests.Wire_IsByteIdentical_AfterInterning`). What is **not** promised is byte identity
  across library *versions*: adding a serialized member to `Workbook` shifts every saved file, because
  MemoryPack writes the member count in the object header. That has now happened twice — `0x01` → `0x02`
  when `DefinedNames` was added (the frozen `workbook-pre-namespaces.msgpack.bin` fixture still carries
  `0x01`) and `0x02` → `0x03` for the table registry (see [Forward-compatibility: the table
  registry](#forward-compatibility-the-table-registry-a-third-workbook-member)). Older *files* keep loading
  in both cases; older *readers* do not.
- **Container** — every other combination is a small self-describing container: the magic `MSWM`, a 1-byte
  format version, the uncompressed model length (int32 LE), then the body. `Load` sniffs the 4-byte magic: a
  match is a container, anything else is a raw (cold or pre-existing) model, so old files keep loading
  unchanged. The version byte selects the body encoding:
  - **v1 (uncompressed warm)** — the **same** model bytes a cold save would write, then a value block (the
    MemoryPack of the cached values). Warm-start files written before compression existed are exactly this.
  - **v2 (Brotli, legacy, read-only)** — the model and value block concatenated and Brotli-compressed as a
    *single* stream, written as two whole-buffer `BrotliStream.Write` calls. Nothing in the library writes
    v2 anymore (superseded by v3, below) — old v2 files still load unchanged, forever, the same one-way
    policy as an append-only union tag.
  - **v3 (Brotli, chunked — the default since the current release)** — the model and value block
    concatenated and Brotli-compressed as a single stream, written as a sequence of **exactly 64KB writes**
    (the final one shorter) to the `BrotliStream` instead of two whole-buffer writes. That fixed chunking is
    part of the format's definition, not an implementation detail: every write mechanism
    (`WorkbookIoBuffering.Pooled` or `.Pipelines`, `Save` or `SaveAsync`) re-chunks whatever bytes it
    produces into the same 64KB windows before they reach Brotli, so all four combinations produce
    byte-identical v3 files — v2 could not make that promise (splitting the same bytes across many small
    `Write` calls measurably changes Brotli's compressed output, so every v2 writer had to fall back to
    materializing the model and value block as two whole buffers to stay byte-identical). See [Measured
    sizes](#measured-sizes) for the compression-ratio win this chunking scheme brought on a large real-world
    workbook.

Because the model and its values travel in one file, they can never desynchronize on load.

### What warm start does *not* freeze

One kind of cached value is deliberately **excluded** from the snapshot and recomputes on first read, even
from a warm file:

- **Volatile cells** — anything that touched `NOW`/`TODAY`/`RAND`/`RANDBETWEEN` (directly or transitively).
  Persisting them would "freeze yesterday's clock"; instead they re-sample on the next read.

The surrogate also refuses a **reference-typed** value, but no cell can produce one any more: the cell
boundary applies
[implicit intersection](workbook-and-expressions.md#implicit-intersection-at-the-cell-boundary) before the
value is stored, so `=MyName` over a range is persisted as its intersected value like any other scalar. The
refusal remains as defence-in-depth over the public `ComputedValue.Reference` API.

### Staleness contract

Warm start persists values you already computed; it does not track edits. The post-load contract is the same
as always: **after editing cells, call `InvalidateCache()`** (or `Recalculate()` for a volatile-only refresh)
before reading, or you will read stale values. A warm load only skips the *first* recomputation of unchanged,
non-volatile cells — it changes nothing about how invalidation works afterwards. And, as with a cold load,
[custom functions](custom-functions.md) must still be re-registered: cells that were **not** cached at save
time (or that you invalidate) will re-evaluate their calls and need the implementation present.

**Excel's date epoch (3.17.0).** The wire format does not change: a date is a `NumberValue` double and stays
one, with no new union tag. What changes is what an early-1900 serial MEANS. From 3.17.0 on, serial 1 is
1900-01-01, serial 0 is Excel's day zero 1900-01-00 and serial 60 is Excel's phantom 1900-02-29; before 3.17.0
the engine read those serials one calendar day earlier (serial 1 was 1899-12-31). Serials from 61 (1900-03-01)
up — every date a real workbook holds — are unaffected. A snapshot written by 3.16.x therefore reloads
byte-identically, but a cached result of a formula over the `[0, 61)` window is now a day off:
`=DATE(1900,1,1)` cached as `2` reads back as `2` until `InvalidateCache()`. If a workbook computes over
early-1900 dates, invalidate the cache once after upgrading.

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

v3's fixed-64KB chunking (see [File format](#file-format)) is not just a determinism/byte-identity
contract — it also compresses *better* than v2's whole-buffer write. On a large, formula-heavy real-world
workbook (~680K cells, `CompressionLevel.Optimal`): v2 produced a 3,616,055-byte file; the same workbook
through v3 produced 3,137,719 bytes — **~13% smaller**, for free, just from writing the compressed stream
in fixed windows instead of one shot.

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

Releases that changed how a saved value is *interpreted* without touching the format:

- **3.17.0, the date epoch** — no format change and no new tag *for that part of the release*. Early-1900
  date serials (`[0, 61)`) change MEANING by one calendar day; cached results over that window need one
  `InvalidateCache()`. Other parts of 3.17.0 do add union tags and a `Workbook` member — see the
  forward-compatibility subsections below.

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

### Forward-compatibility: the `AGGREGATE` node (tag 322)

`AGGREGATE` is a new expression node type and claims the next append-only union tag, **322** (see
[Function reference](function-reference.md) for what the function does). A cell whose formula calls it is
serialized under that tag.

This is a **one-way** compatibility boundary, same as any append-only tag addition:

- A file saved by **this or a later** version of the library — whether produced by `Workbook.Save` or by
  `ExcelFile.Load` followed by a save — can contain cells using tag 322 whenever a cell's formula is an
  `AGGREGATE` call. Such a file **cannot be opened by a version of the library older than the one that
  introduced this tag**: the older MemoryPack union does not recognize it and deserialization fails.
- A file saved by an **older** version of the library never contains this tag, and continues to load
  unchanged in this and every later version, exactly as the append-only policy above guarantees.

**What a new tag does *not* break.** The tag is written per node, not per file, so a workbook that does not
use `AGGREGATE` serializes to exactly the same bytes as before: the frozen binary goldens in the test suite
— the base64 cell-store wire snapshot and the pre-2.0 `.msgpack.bin` fixture — stay valid and need no
regeneration. Only a new **member on `Workbook` itself** would change the shape of every saved file and
force that.

### Forward-compatibility: the dynamic-array producer nodes (tags 323-326)

`FILTER`, `SORT`, `UNIQUE` and `SEQUENCE` are four new expression node types and claim the next four
append-only union tags, in one coordinated edit (see [implicit array
arguments](workbook-and-expressions.md#dynamic-array-producers) for what they do):

| Tag | Node |
| --: | --- |
| 323 | `Lookup.Filter` |
| 324 | `Lookup.Sort` |
| 325 | `Lookup.Unique` |
| 326 | `Mathematics.Sequence` |

A cell whose formula calls one of them is serialized under that tag. The next free tag is **327**.

This is a **one-way** compatibility boundary, same as any append-only tag addition:

- A file saved by **this or a later** version of the library — whether produced by `Workbook.Save` or by
  `ExcelFile.Load` followed by a save — can contain cells using tags 323-326 whenever a cell's formula calls
  one of the four. Such a file **cannot be opened by a version of the library older than the one that
  introduced these tags**: the older MemoryPack union does not recognize them and deserialization fails.
- A file saved by an **older** version of the library never contains these tags, and continues to load
  unchanged in this and every later version, exactly as the append-only policy above guarantees.
- As with `AGGREGATE`, the tags are written per node, so a workbook that uses none of the four serializes to
  exactly the same bytes as before and the frozen binary goldens need no regeneration.

**A warm snapshot can now carry error code 7, `#CALC!`.** These four introduce Excel's empty-array error as
the eighth [`Error`](computed-value.md) code, so a warm-start value block can hold a `CachedCellValue` whose
`ErrorCode` is `7` — for example the cached result of `=SUM(FILTER(A1:A3,A1:A3>100))`. That is a **value**,
not a format change: the value block's shape is unchanged and no tag is involved.

The degradation is graceful in both directions. What is *pinned by tests* is the mechanism, on this build:

- A code past the end of the error table displays as the unknown marker **`#ERR?`** rather than throwing —
  `Error.FromCode(8).Display` is `#ERR?` — which is the rule a **pre-3.17** build applies to code 7, whose
  table stopped at `#N/A`. So an older build reading such a snapshot shows a wrong error text; it does not
  crash and it does not fail to load. (The old build itself is not exercised by the suite; what the suite
  pins is that the marker rule exists and that 7 is no longer past the table.)
- An error display text the engine does not know folds onto `#VALUE!` rather than throwing —
  `Error.FromDisplay("#SPILL!")` is `#VALUE!` — which is also what makes `#CALC!` demonstrably a *real* code
  now rather than one the fold conjures.
- The warm-start surrogate round-trips code 7 exactly: `CachedCellValue.ErrorCode` is `7` going out and
  `Error.Calc` coming back.

Those three are pinned by `ErrorTests.Calc_IsTheEighthError_AndRoundTripsExactly` and
`ErrorTests.AnUnknownDisplay_StillFoldsOntoValue_AndCalcIsNoLongerUnknown`. The `.xlsx` round trip also
closes — the exporter writes `#CALC!` as an ordinary `t="e"` cell and the loader recognizes it alongside the
other singleton codes — but it is carried by the code (the two mirrored canonicalization switches) rather
than by a dedicated fixture in the Excel suite.

### Forward-compatibility: container v3 (chunked Brotli)

`WorkbookCompression.Brotli` now writes container version 3 by default (see [File
format](#file-format)) instead of v2. This is also a **one-way** boundary:

- A file saved by **this or a later** version of the library with `Compression = Brotli` is a v3 container.
  Such a file **cannot be opened by a version of the library older than the one that introduced v3**: the
  older reader's container-version switch does not recognize tag 3 and `Load` throws `InvalidDataException`.
- A file saved by an **older** version of the library (v1 or v2) continues to load unchanged in this and
  every later version — `Load` still recognizes and decompresses v2 containers, it just never writes them.

There is no option to opt back into writing v2 — the same append-only-in-spirit policy as the tags above:
a superseded write format is dropped, not kept as a knob, while the reader keeps it forever.

### Forward-compatibility: the table registry (a third `Workbook` member)

`Workbook` now serializes a **third** member: the table registry that `Workbook.DefineTable` writes and
`Workbook.Tables` exposes (see [Workbook, sheets and expressions →
Tables](workbook-and-expressions.md#tables)). Unlike a new union tag, a new member on `Workbook` itself
changes the shape of **every** saved file — the object header carries the member count — so this boundary
applies whether or not the workbook holds a single table, and whether the file came from
`Workbook.Save`/`SaveAsync` or from a save after `ExcelFile.Load`. (`ExcelFile.Load` does not populate the
registry from an xlsx `<table>` part yet — see [Excel interop → Scope and
limitations](excel-interop.md#scope-and-limitations) — but a workbook it produced still gets the new header
when you save it.)

This is a **one-way** compatibility boundary:

- **No new union tag.** The registry lives on `Workbook`, not in the expression union, and nothing in the
  formula language reads a table yet: a structured reference (`Tabela1[Valor]`) does not parse. The node
  that will represent one, and the union tag it claims (the next free tag is **327** — 0-326 are taken; count
  `Expression.cs` with a leading-bracket anchor, because a bare `grep -c MemoryPackUnion` is one too high),
  belong to the reference-semantics work; that half of the boundary does not exist yet.
- **The object header goes `0x02` → `0x03`, and an empty registry costs four bytes.** MEASURED 2026-09-10
  on the branch that introduces the registry (MemoryPack 1.21.4, cold uncompressed `Save`): an empty
  `Workbook` serializes to **13** bytes (`03` + three zero-length maps) where the previous version wrote
  **9** (`02` + two); byte 0 is the only byte that changes, and the four added bytes are the empty registry
  at the very end. On a populated workbook the frozen cell-store golden goes **726 → 730** bytes with every
  byte after the header still at its old offset — pinned by
  `CellStoreTests.Wire_NewGolden_IsPreTablesGoldenPlusEmptyTablesMember` and its twin in
  `SheetNameInterningTests` (429 → 433 bytes).
- **A file saved by this or a later version cannot be opened by an older one.** The older `Workbook`
  declares two members and MemoryPack refuses a header that announces three. MEASURED 2026-09-10, an
  older-version reader over a file this version wrote (both a raw `MemoryPackSerializer.Deserialize` and
  `Workbook.Load` throw it, and both for an empty workbook and one holding a table):
  `MemoryPackSerializationException: Danfma.MySheet.Workbook property count is 2 but binary's header maked
  as 3, can't deserialize about versioning.` — "maked" is MemoryPack's own typo, and searching for it
  verbatim is what should lead a user to this subsection. Same one-way shape as an appended union tag, but
  it applies to **every** file, not only to files that use the new feature.
- **Files saved by an older version keep loading — forever, with an empty registry.** MemoryPack tolerates
  a header count *below* the declared member count and leaves the absent members `null`, and `Workbook`'s
  `[MemoryPackOnDeserialized]` hook turns that `null` into an empty (case-insensitive) registry. The
  precedent is already frozen in the suite: `workbook-pre-namespaces.msgpack.bin` begins `01 02 00 00 00 …`
  — a ONE-member `Workbook`, written before `DefinedNames` existed — and still loads
  (`MemoryPackCompatibilityTests.PreNamespaceFixture_LoadsAndReevaluates`), as do the two-member (`0x02`)
  cell-store golden (`CellStoreTests.PreTablesGolden_StillLoads_WithEmptyTables`) and the v2 warm container
  whose compressed body is a two-member model
  (`ContainerVersionCompatibilityTests.GoldenV2Fixture_IsVersion2_AndLoadsForever`) — all three arrive with
  `Tables.Count == 0`.

**What this breaks that a new tag does not.** The `AGGREGATE` note above ends by saying only a new member on
`Workbook` itself would change the shape of every saved file and force the frozen goldens to be
regenerated. This is that case: both wire goldens were regenerated for it, and *mechanically* — the new
constant is `[0x03] + old[1..] + four zero bytes`, never a captured debug print — so the constants
themselves record exactly what moved.

## When to use which format

| Need | Use |
| --- | --- |
| Fast native persistence of a computed model (cache-style, service restarts, snapshots between processing steps) | `Workbook.Save` / `Load` |
| Interchange with people or other tools (open in Excel, send a report) | [`SaveAsExcel` / `MergeIntoExcel`](excel-interop.md) |
| Ingesting the source-of-truth spreadsheet | [`ExcelFile.Load`](excel-interop.md) |
