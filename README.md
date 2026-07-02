# Danfma.MySheet

[![NuGet](https://img.shields.io/nuget/v/Danfma.MySheet.svg?label=Danfma.MySheet)](https://www.nuget.org/packages/Danfma.MySheet/)
[![NuGet](https://img.shields.io/nuget/v/Danfma.MySheet.Excel.svg?label=Danfma.MySheet.Excel)](https://www.nuget.org/packages/Danfma.MySheet.Excel/)

> 🇧🇷 Documentação em português: [docs/pt-BR](docs/pt-BR/README.md)

A fast, in-memory spreadsheet formula engine for .NET — parse Excel-style formulas, evaluate them, and
extract data, without a spreadsheet application.

## Why MySheet

MySheet does not try to compete with ClosedXML, EPPlus, NPOI or the other established Excel libraries —
they are very good at what they do, which is reading and writing Excel files with high fidelity. MySheet
solves a narrower problem, in a simpler and faster way:

> You keep an Excel workbook on a server as the **source of truth for a calculation**. You need to load
> it, re-evaluate its formulas with fresh inputs, and expose or write back the results — repeatedly, with
> low overhead, and without Excel installed.

For that scenario, MySheet gives you:

- **A real formula engine.** Formulas are parsed into expression trees and re-evaluated in-process by
  MySheet itself — no Excel installation, no COM automation, no shelling out.
- **A deliberately small API.** A `Workbook` holds `Sheet`s, a `Sheet` holds `Expression`s, and
  `Evaluate` returns a `ComputedValue`. That is most of the surface.
- **A performance-first design.** Evaluating numbers, booleans, blanks and errors allocates nothing
  (`ComputedValue` is a value-type union), cell results are memoized, and deep dependency chains run on a
  dedicated large-stack thread instead of overflowing.

And to be equally honest about scope: MySheet implements **155 built-in functions** (plus your own custom
ones) out of Excel's ~520, and the Excel interop intentionally skips styles, number formats and other
presentation features in its current MVP. If you need full-fidelity spreadsheet manipulation, the
libraries above remain the right tools — they also combine well with MySheet (the test suite itself uses
ClosedXML as an independent oracle).

## Packages

| Package | What it does |
| --- | --- |
| [`Danfma.MySheet`](https://www.nuget.org/packages/Danfma.MySheet/) | The core engine: formula parser, 155 built-in functions, custom functions, allocation-free evaluation, per-cell memoization, MemoryPack serialization. |
| [`Danfma.MySheet.Excel`](https://www.nuget.org/packages/Danfma.MySheet.Excel/) | Excel (`.xlsx`) interop via the OpenXML SDK: load workbooks (formulas become real expression trees), export to `.xlsx`, and merge computed values into existing templates. |

The two packages are released in lockstep and always share the same version.

```shell
dotnet add package Danfma.MySheet
dotnet add package Danfma.MySheet.Excel   # only if you need .xlsx interop
```

Both target **.NET 10**.

## Quick start

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

sheet["A1"] = new NumberValue(1);
sheet["A2"] = new NumberValue(2);
sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);

// Typed, allocation-free result (memoized per cell):
ComputedValue result = workbook.GetCellValue("Sheet1", "A3");

double total = result.ToDouble();            // 3.0 (throws if not a number)

if (result.TryGetError(out Error error))     // e.g. error.Display == "#DIV/0!"
{
    // handle the error
}

object? boxed = result.AsObject();           // 3.0 (double) — for object?-based interop
```

Excel as the source of truth, end to end:

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Excel;
using Danfma.MySheet.Expressions;

// Load: formulas come back as real expression trees, re-evaluated by MySheet.
Workbook workbook = ExcelFile.Load("model.xlsx");

// Feed fresh inputs and read the recalculated outputs.
var inputs = workbook["Inputs"];
inputs["B1"] = new NumberValue(1250);
workbook.InvalidateCache();

double projection = workbook.GetCellValue("Results", "B10").ToDouble();

// Produce a report from a pristine template: copy it, then merge the computed
// values into the copy (formatting preserved, formulas replaced by literal values).
File.Copy("template.xlsx", "report.xlsx", overwrite: true);
workbook.MergeIntoExcel("report.xlsx");
```

## Features

- **Formula parser** (Pratt parser) for the Excel operator set: `+ - * / ^ %`, `&` (text), comparisons
  `= <> < > <= >=` (Excel cross-type ordering), and references `: ! ,` plus grouping `( )`.
- **155 built-in functions**: logical (`IF/IFS/SWITCH/AND/OR/XOR/NOT/IFERROR/IFNA/LET/TRUE/FALSE`),
  aggregation (`SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/COUNTBLANK`), conditional aggregation
  (`COUNTIF(S)/SUMIF(S)`), text (`UPPER/LOWER/TRIM/LEN/LEFT/MID/RIGHT`, search & replace
  `FIND/SEARCH/REPLACE/SUBSTITUTE/TEXTBEFORE/TEXTAFTER` plus `REGEXTEST/REGEXEXTRACT/REGEXREPLACE`,
  codes `CHAR/CODE/UNICHAR/UNICODE/CLEAN`, formatting `TEXT/FIXED/DOLLAR/VALUE/NUMBERVALUE/VALUETOTEXT`,
  `CONCAT/CONCATENATE/TEXTJOIN/REPT/PROPER/EXACT/T`), math & trigonometry (`SQRT/POWER/EXP/LN/LOG`,
  rounding `TRUNC/MROUND/CEILING/FLOOR` and variants, `MOD/QUOTIENT/PRODUCT`, combinatorics
  `FACT/COMBIN/GCD/LCM`, the full trig/hyperbolic set, `ROMAN/ARABIC/BASE/DECIMAL`), info (the `IS*`
  family, `N/NA/TYPE/ERROR.TYPE/SHEET/SHEETS`), lookup (`ROW/ROWS/MATCH/INDEX/VLOOKUP/XLOOKUP/OFFSET`),
  and financial (`PMT/PV/FV/NPER/IPMT/PPMT/NPV/RATE/IRR`).
- **References**: sheet-qualified (`Sheet2!A1`, `'My Sheet'!A1:B2`), absolute markers (`$A$1`),
  reference unions (`(A1:A3, C1:C3)`), and case-insensitive sheet names.
- **Custom functions**: register host functions by name (`workbook.RegisterFunction`) — arguments arrive
  unevaluated (lazy, short-circuit friendly), and calls parse and serialize with the workbook.
- **Allocation-free evaluation**: `expression.Evaluate(workbook)` returns a `ComputedValue` — an opaque
  value-type union (number / boolean / text / blank / error / reference) that does **not** box numbers.
  Extract strictly with `TryGetNumber`/`AsDouble`/`ToDouble` (and `TryGetError(out Error)`), or get the
  boxed `object?` via `AsObject()` for interop.
- **Memoization**: per-cell cache (storing `ComputedValue` inline — no long-lived per-cell box) with
  explicit invalidation (`InvalidateCache`); circular references become `#REF!` instead of a stack
  overflow.
- **MemoryPack serialization** of the workbook (`Workbook.Save`/`Load`, sync and async).
- **Excel (.xlsx) interop** via `Danfma.MySheet.Excel` (OpenXML SDK, cross-platform, no Excel
  installation): `ExcelFile.Load`, `SaveAsExcel` (`ValuesOnly` snapshot or `Formulas` with cached
  values), and `MergeIntoExcel` (inject computed values into an existing file, preserving formatting).

## Documentation

| Guide | What it covers |
| --- | --- |
| [Getting started](docs/getting-started.md) | Install, build a workbook, parse and evaluate your first formulas. |
| [Workbook, sheets and expressions](docs/workbook-and-expressions.md) | The object model, parsing rules, operators, references, and turning expressions back into formula text. |
| [ComputedValue and errors](docs/computed-value.md) | The result type: kinds, strict extraction (`TryGet*`/`As*`/`To*`), `AsObject()`, and the `Error` struct. |
| [Custom functions](docs/custom-functions.md) | Registering host functions, lazy arguments, implicit return conversions, ranges, and serialization/re-registration. |
| [Excel interop](docs/excel-interop.md) | `ExcelFile.Load`, `SaveAsExcel` and `FormulaMode`, `MergeIntoExcel`, the template→report recipe, and scope limits. |
| [Serialization](docs/serialization.md) | MemoryPack `Save`/`Load`, what round-trips, and what must be re-registered. |
| [Performance](docs/performance.md) | Memoization, cache invalidation, `RunWithLargeStack`, and the allocation-free design (with measured numbers). |
| [Function reference](docs/function-reference.md) | All 155 built-in functions by category, plus the full Excel coverage table. |

## Excel function coverage

MySheet implements 155 of the ~520 functions in Microsoft's official Excel function catalog. The full
per-category coverage table (implemented vs. not yet) lives in the
[function reference](docs/function-reference.md#excel-function-coverage); the authoritative registered
list is the `Functions` map in [`Danfma.MySheet/Parsing/Parser.cs`](Danfma.MySheet/Parsing/Parser.cs).

## License

MIT — see [LICENSE](LICENSE).
