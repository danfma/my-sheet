# Danfma.MySheet

A fast, in-memory spreadsheet formula engine for .NET — parse Excel-style formulas, evaluate them, and
extract data, without a full spreadsheet application.

## Features

- **Formula parser** (Pratt parser) for the Excel operator set: `+ - * / ^ %`, `&` (text), comparisons
  `= <> < > <= >=` (Excel cross-type ordering), and references `: ! ,` plus grouping `( )`.
- **~40 built-in functions**: `IF/AND/OR/NOT/IFERROR/IFNA`, `SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/COUNTBLANK`,
  `COUNTIF(S)/SUMIF(S)`, text (`UPPER/LOWER/TRIM/LEN/LEFT/MID/VALUE/CONCAT/CONCATENATE/TEXTJOIN/TEXT`),
  math (`INT/ROUND/ROUNDUP/ABS`), info (`ISNUMBER/ISBLANK`), lookup (`ROW/ROWS/SHEET/MATCH/INDEX/VLOOKUP/
  XLOOKUP/OFFSET`), and `LET`.
- **References**: sheet-qualified (`Sheet2!A1`, `'My Sheet'!A1:B2`), absolute markers (`$A$1`), and
  case-insensitive sheet names.
- **Custom functions**: register host functions by name (`Workbook.RegisterFunction`) — they parse and
  serialize with the workbook.
- **Memoization**: per-cell value cache with explicit invalidation; circular references become `#REF!`
  instead of a stack overflow.
- **MemoryPack serialization** of the workbook.

## Quick start

```
dotnet add package Danfma.MySheet
```

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Parsing;
using Danfma.MySheet.Expressions;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");

sheet["A1"] = new NumberValue(1);
sheet["A2"] = new NumberValue(2);
sheet["A3"] = ExpressionParser.Parse("=SUM(A1:A2)", sheet);

var total = sheet["A3"].Compute(workbook); // 3.0
```

For deep dependency chains, wrap the evaluation in a large-stack thread:

```csharp
var value = Workbook.RunWithLargeStack(() => sheet["A3"].Compute(workbook));
```

## License

MIT — see [LICENSE](LICENSE).
