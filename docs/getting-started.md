# Getting started

MySheet is an in-memory spreadsheet formula engine. You build (or load) a `Workbook`, put values and
formulas into its `Sheet`s, and ask the engine for computed results — no Excel installation, no COM, and
no spreadsheet UI involved.

## Install

```shell
dotnet add package Danfma.MySheet
dotnet add package Danfma.MySheet.Excel   # only if you need .xlsx interop
```

Both packages target **.NET 10** and are always released together with the same version.

## Your first workbook

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

// Literal values are Expression nodes too:
sheet["A1"] = new NumberValue(10);
sheet["A2"] = new NumberValue(32);

// Formulas are parsed into real expression trees:
sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);
sheet["A4"] = ExpressionParser.Parse("=IF(A3>40, \"big\", \"small\")", sheet);
```

`ExpressionParser.Parse` follows the cell-entry convention you know from Excel: a string starting with
`=` is a formula; anything else is a literal (number, boolean, or text). See
[Workbook, sheets and expressions](workbook-and-expressions.md) for the full rules.

## Evaluating

The one evaluation API is `Evaluate`, which returns a [`ComputedValue`](computed-value.md) — a value-type
union of number / boolean / text / blank / error / reference that never boxes numbers:

```csharp
ComputedValue total = sheet["A3"].Evaluate(workbook);
ComputedValue label = sheet["A4"].Evaluate(workbook);

double sum = total.ToDouble();     // 42.0 — strict: throws if the result is not a number
string text = label.ToText();      // "big"
```

For cell reads, prefer `Workbook.GetCellValue` — it returns the same `ComputedValue` but memoizes it per
cell, so a cell referenced by many formulas is computed once:

```csharp
ComputedValue cached = workbook.GetCellValue("Sheet1", "A3");
```

Extraction is explicit and strict — there is no hidden coercion at the API surface:

```csharp
var value = workbook.GetCellValue("Sheet1", "A3");

if (value.TryGetNumber(out var number))
{
    Console.WriteLine($"Number: {number}");
}
else if (value.TryGetError(out var error))
{
    Console.WriteLine($"Formula failed: {error}");   // prints e.g. "#DIV/0!"
}

double? maybe = value.AsDouble();   // null when the kind does not match
object? boxed = value.AsObject();   // bridge to object?-based code (boxes numbers)
```

## Editing cells and the cache

The memoization cache is **not** invalidated automatically. After editing cells, call
`InvalidateCache()` before reading again:

```csharp
sheet["A1"] = new NumberValue(100);
workbook.InvalidateCache();

double updated = workbook.GetCellValue("Sheet1", "A3").ToDouble();   // 132.0
```

This is a deliberate design for the primary use case — read-heavy extraction with rare, batched
mutations. See [Performance](performance.md) for the rationale.

## Deep dependency chains

Evaluation is recursive. If your workbook has very long dependency chains (for example, a cumulative
column thousands of rows deep), wrap the whole extraction batch in `RunWithLargeStack`, which runs the
work on a thread with a large reserved stack:

```csharp
var value = Workbook.RunWithLargeStack(() => workbook.GetCellValue("Sheet1", "A20000"));
```

## Loading an Excel file

With the `Danfma.MySheet.Excel` package, an `.xlsx` file becomes a regular `Workbook` whose formulas are
real expression trees, re-evaluated by MySheet:

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Excel;

Workbook workbook = ExcelFile.Load("model.xlsx");
double result = workbook.GetCellValue("Data", "B10").ToDouble();
```

See [Excel interop](excel-interop.md) for exporting workbooks and merging computed values into templates.

## Where to go next

- [Workbook, sheets and expressions](workbook-and-expressions.md) — the object model in depth.
- [ComputedValue and errors](computed-value.md) — everything about reading results.
- [Custom functions](custom-functions.md) — add your own functions to the engine.
- [Function reference](function-reference.md) — the 112 built-in functions.
