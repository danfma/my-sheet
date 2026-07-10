# Excel interop (`Danfma.MySheet.Excel`)

`Danfma.MySheet.Excel` connects the MySheet engine to real `.xlsx` files through the OpenXML SDK —
cross-platform, **no Excel installation, no COM**. This is the package that enables MySheet's core
scenario: keep an Excel workbook on a server as the source of truth, load it, re-evaluate its formulas
in-process, and expose or write back the results.

It is deliberately *not* a general-purpose Excel manipulation library. It moves **values and formulas**
between `.xlsx` and the engine; styles, number formats and other presentation features are out of scope
in the current MVP (see [Scope and limitations](#scope-and-limitations)). If you need those, ClosedXML,
EPPlus or NPOI remain excellent choices — and combine well with MySheet.

```shell
dotnet add package Danfma.MySheet.Excel
```

## Loading: `ExcelFile.Load`

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Excel;

Workbook workbook = ExcelFile.Load("model.xlsx");

// or from any readable stream:
using var stream = File.OpenRead("model.xlsx");
Workbook fromStream = ExcelFile.Load(stream);
```

Loading is **fully streaming**: shared strings and worksheets are read with a forward-only `XmlReader`,
so the OpenXML DOM is never materialized and the only full in-memory representation is the MySheet model
itself (on a ~620k-cell file: ~4.8x faster and ~6.6x fewer allocations than a DOM-based load). Two
practical notes: the `Stream` overload requires a readable **and seekable** stream (a requirement of the
underlying package reader), and cells or rows without an `r` attribute — implicit positions, emitted by
some minimal writers — load at their spec-defined position.

The result is a plain MySheet `Workbook`. The key property: **formula cells become real `Expression`
trees**, parsed by the MySheet parser and re-evaluated by the MySheet engine — the cached values stored
in the file are ignored for formula cells. Change an input cell, invalidate the cache, and every
dependent formula recomputes:

```csharp
using Danfma.MySheet.Expressions;

var workbook = ExcelFile.Load("model.xlsx");

workbook["Inputs"]["B1"] = new NumberValue(2500);
workbook.InvalidateCache();

double updated = workbook.GetCellValue("Results", "B10").ToDouble();
```

How file content maps into the workbook:

| In the `.xlsx` | In the `Workbook` |
| --- | --- |
| Formula cell (`<f>`) | Parsed `Expression` tree (re-evaluated by MySheet; the cached `<v>` is ignored). |
| Number cell | `NumberValue`. |
| Shared / inline string | `StringValue` (rich-text runs flattened to plain text). |
| Boolean cell | `BooleanValue`. |
| Error cell | `ErrorValue` (evaluates to the corresponding `Error`). |
| Date cell | `NumberValue` with the Excel **serial number** (ISO-8601 dates in strict-mode files are converted via `ToOADate`). |
| Style-only / empty cell | Nothing stored — reads as blank. |
| Shared-formula "slave" (a dragged formula cell carrying no formula text) | Expanded into a real formula: the master's text is shifted by the row/column delta (relative references move, `$`-anchored components stay, text inside string literals is untouched) and parsed like any other formula. |
| Workbook-scoped defined name (`<definedName>`) | An entry in [`Workbook.DefinedNames`](workbook-and-expressions.md#named-ranges): the `refersTo` text is parsed as a formula. **Sheet-scoped** names (those with a `localSheetId`) and Excel's **builtin `_xlnm.*`** names (`Print_Area`, `Print_Titles`, `_FilterDatabase`, …) are skipped. |

Sheets are created in the file's tab order, so `Sheet.Index` (and the `SHEET` function) match Excel.

If the file uses functions MySheet does not implement, those cells parse into `FunctionCall` nodes and
evaluate to `#NAME?` — unless you provide the behavior yourself via
[`RegisterFunction`](custom-functions.md), which is the intended escape hatch.

A handful of load-time issues (an invalid defined name, a date literal that fails to parse) are skipped
rather than failing the whole load — by default silently, exactly as before. Pass `ExcelLoadOptions` with
an `OnWarning` callback to `Load` to observe them instead:

```csharp
var warnings = new List<ExcelLoadWarning>();
var workbook = ExcelFile.Load("model.xlsx", new ExcelLoadOptions { OnWarning = warnings.Add });
```

Each `ExcelLoadWarning` carries a `Kind` (`InvalidDefinedName` or `UnparsableDateLiteral`), a `Subject`
(the defined name, or the cell id) and a `Detail` string. The callback is a simple `Action<T>` rather than
an accumulated list, so the host decides whether to log, collect, or ignore each warning.

## Exporting: `SaveAsExcel`

```csharp
using Danfma.MySheet.Excel;

// Default: a flattened snapshot — every cell written as its computed literal value.
workbook.SaveAsExcel("snapshot.xlsx");

// Keep formulas alive: write the formula text plus its computed (cached) value.
workbook.SaveAsExcel("live.xlsx", new ExcelExportOptions { FormulaMode = FormulaMode.Formulas });

// Stream overload:
using var output = File.Create("snapshot.xlsx");
workbook.SaveAsExcel(output);
```

`SaveAsExcel` creates a **new** file containing the workbook's sheets and cells. Before writing, every
cell is evaluated up front on a single large-stack thread (`RunWithLargeStack`) with memoization, so
deep dependency chains cannot overflow mid-write.

`ExcelExportOptions.FormulaMode` controls formula cells:

| Mode | Formula cells become | Use when |
| --- | --- | --- |
| `FormulaMode.ValuesOnly` (default) | Their computed literal value — a flattened snapshot with no formulas. | The recipient should see results, not logic (reports, data handoff). |
| `FormulaMode.Formulas` | The Excel formula (`<f>`, rendered by `FormulaWriter`) **plus** its computed value as the cached `<v>`. | The file should keep recalculating when opened in Excel. |

Details worth knowing:

- Blank results are omitted entirely (like Excel's own files).
- Text literals are deduplicated through a shared-string table; text produced *by a formula* is written
  as the formula's cached string, per the `.xlsx` convention.
- A cell whose result is a bare reference (e.g. a multi-cell `OFFSET` used as a scalar) is written as
  `#VALUE!`, matching how the engine treats it.
- In `Formulas` mode, calls to [custom functions](custom-functions.md) are written with their registered
  name — Excel will show the cached value and flag the unknown function, which is expected.
- [Named ranges](workbook-and-expressions.md#named-ranges) in `Workbook.DefinedNames` are written as
  workbook-scoped `<definedName>` entries, in **both** formula modes. The `refersTo` text is fully
  qualified (`FormulaWriter` with an empty context, so every reference keeps its `Sheet!` prefix).

## Merging into a template: `MergeIntoExcel`

```csharp
workbook.MergeIntoExcel("report.xlsx");
```

`MergeIntoExcel` edits an **existing** file **in place**: every cell held by the MySheet workbook is
written into the target as its computed literal value, while *everything else* in the file — styles,
number formats, other cells, other sheets, charts — is left untouched. This is the tool for the classic
"beautiful template, computed numbers" report flow:

```csharp
// The template→report recipe: copy the pristine template, merge into the copy.
File.Copy("template.xlsx", "report.xlsx", overwrite: true);
workbook.MergeIntoExcel("report.xlsx");
```

The in-place-only design is deliberate: merging *mutates the file you give it*, and creating files is
`SaveAsExcel`'s job. There is intentionally no `MergeIntoExcel(template, output)` overload — the
`File.Copy` step keeps the template pristine and makes the mutation explicit.

Merge semantics:

- Sheets are matched **by name, case-insensitively**. Workbook sheets missing from the target are
  skipped.
- Each written cell gets the **literal computed value** — any formula the target cell had is dropped
  (the merged file shows your engine's numbers, not Excel's recalculation).
- **Blank values are not written**, leaving the target cell exactly as it was.
- Cell formatting is preserved: only the content is replaced; the cell's style reference is untouched.
- Text is written as an inline string, so the target's shared-string table is not modified.
- Missing rows/cells are created in the correct OpenXML order as needed.
- As with `SaveAsExcel`, all values are computed up front via `RunWithLargeStack` with memoization.

## Scope and limitations

Being honest about what the interop MVP does **not** do:

- **No styles or presentation**: fonts, colors, number formats, column widths, merged cells, charts and
  the like are not modeled. `Load` ignores them; `SaveAsExcel` does not produce them; `MergeIntoExcel`
  *preserves* the target's existing formatting but cannot create it.
- **Dates are serial numbers**: they enter and leave as `double`s (Excel's own representation). Apply
  date formatting in the template (merge flow) or convert with `DateTime.FromOADate` in your code.
- **Absolute markers are not preserved on write**: `$A$1` parses fine (it identifies the same cell) but
  un-parses as `A1` — a fidelity loss only in `FormulaMode.Formulas` exports, and only cosmetic unless
  you plan to copy/fill formulas in Excel afterwards.
- **Shared formulas ARE expanded**: slave cells of a dragged formula (which carry no formula text in the
  file) are rebuilt from the master's text, shifting relative references by the cell delta while keeping
  `$`-anchored components fixed — so they stay real formulas that react to input changes. A slave whose
  master is missing from the file falls back to its cached literal value.
- **Function coverage is not total** (see the current count and list of built-ins, plus your custom functions, in the
  [function reference](function-reference.md). Formulas using other functions load as `FunctionCall`
  nodes and evaluate to `#NAME?` unless registered.
- **Only workbook-scoped defined names cross over**: sheet-scoped names (with a `localSheetId`) and the
  builtin `_xlnm.*` names (print areas, filter databases, …) are skipped on load, and MySheet only ever
  writes workbook-scoped names. A defined name whose `refersTo` cannot be parsed is skipped rather than
  failing the load.

## See also

- [Getting started](getting-started.md) — the end-to-end flow in miniature.
- [Custom functions](custom-functions.md) — supplying behavior for functions the engine lacks.
- [Performance](performance.md) — why exports evaluate through `RunWithLargeStack`.
