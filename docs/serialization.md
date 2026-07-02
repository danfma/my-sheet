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

| | Persisted? | Notes |
| --- | --- | --- |
| Sheets (name, tab order) | Yes | The case-insensitive name lookup is restored on deserialization. |
| Cells and full expression trees | Yes | Formulas stay formulas — a loaded workbook keeps recalculating. |
| Custom-function **calls** (`FunctionCall` nodes) | Yes | Name and argument expressions round-trip. |
| Custom-function **implementations** (delegates) | **No** | Behavior is code, not data — re-register after loading. |
| Memoization cache | **No** | Values are recomputed lazily on first read after loading. |

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
