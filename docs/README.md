# MySheet documentation

> 🇧🇷 Documentação em português: [pt-BR](pt-BR/README.md)

MySheet is a fast, in-memory spreadsheet formula engine for .NET. It is not a replacement for
full-featured Excel libraries (ClosedXML, EPPlus, NPOI, …) — it is a simpler, faster option for one
specific job: keeping an Excel workbook on a server as the **source of truth for a calculation**, loading
it, re-evaluating its formulas in-process (no Excel installed, no COM, minimal allocation), and exposing
or writing back the results.

## Guides

1. [Getting started](getting-started.md) — install, build a workbook, parse and evaluate formulas.
2. [Workbook, sheets and expressions](workbook-and-expressions.md) — the object model, parsing rules,
   operators, references, and un-parsing back to formula text.
3. [ComputedValue and errors](computed-value.md) — the evaluation result type and the `Error` struct.
4. [Custom functions](custom-functions.md) — extend the engine with your own functions.
5. [Excel interop](excel-interop.md) — load `.xlsx` files, export, and merge into templates.
6. [Serialization](serialization.md) — MemoryPack `Save`/`Load` and what round-trips.
7. [Performance](performance.md) — memoization, `RunWithLargeStack`, and the allocation-free design.
8. [Function reference](function-reference.md) — every built-in function, plus the Excel coverage table.
9. [Migrating to 2.0](migrating-to-2.0.md) — the 2.0 namespace reorganization: type → namespace map,
   before/after `using` examples, and the serialization-compatibility guarantee.
10. [Migrating to 3.0](migrating-to-3.0.md) — the read-only `Sheet.Cells` view, the new `Remove`, and
    the write-maintained structural index (serialization unchanged).

## Packages

| Package | Contents |
| --- | --- |
| `Danfma.MySheet` | Core engine: parser, evaluator, built-in and custom functions, memoization, MemoryPack serialization. |
| `Danfma.MySheet.Excel` | `.xlsx` interop (OpenXML SDK): `ExcelFile.Load`, `SaveAsExcel`, `MergeIntoExcel`. |

Both packages target .NET 10 and are released in lockstep with matching versions.
