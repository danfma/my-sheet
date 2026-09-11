# Workbook, sheets and expressions

This guide covers MySheet's object model — `Workbook`, `Sheet`, and the `Expression` tree — plus the
parsing rules, the operator set, references, and how to turn an expression back into formula text.

## Workbook

A `Workbook` is the root object: a set of named sheets plus the evaluation services (memoization cache,
custom-function registry, serialization).

```csharp
using Danfma.MySheet;

var workbook = new Workbook();

var sheet = workbook.Sheets.Add("Sheet1");   // creates (or returns) a sheet by name
var same = workbook["Sheet1"];               // indexer access
```

- **Sheet names are case-insensitive**, like in Excel: `workbook["sheet1"]` and `workbook["SHEET1"]`
  reach the same sheet.
- The `this[string]` indexer **throws `KeyNotFoundException`** for a name that has no sheet — it is a
  direct host lookup, like a dictionary. Use `TryGetSheet(name, out sheet)` to probe without a
  `try`/`catch`. (A missing sheet referenced *inside a formula* is a different story: it resolves to
  `#REF!` rather than throwing — see [evaluation error semantics](#parsing).)
- `Sheets` is a `ConcurrentDictionary<string, Sheet>`, safe for concurrent readers (the intended
  background-extraction scenario).
- `Sheets.Add(name)` assigns the sheet an `Index` equal to its insertion order — this is what the
  `SHEET` function reports, and what defines tab order when exporting to Excel.

Key `Workbook` members:

| Member | Purpose |
| --- | --- |
| `Sheets` / `this[string]` | Access sheets by name (case-insensitive); the indexer throws `KeyNotFoundException` for a missing name. |
| `TryGetSheet(name, out sheet)` | Non-throwing sheet lookup (case-insensitive) → `bool`; the host's counterpart to the throwing indexer. |
| `GetCellValue(sheetName, id)` | Memoized evaluation of one cell → `ComputedValue`; a reference to a missing sheet resolves to `#REF!` (never throws). |
| `Sheet.CellAddresses` / `Sheet.EnumerateCells()` | Populated-cell enumeration (struct enumerators, no interface boxing): `CellAddresses` yields `(Column, Row)` allocation-free (canonical cells only); `EnumerateCells()` yields `(Id, Column, Row)` deriving the canonical id once per cell (overflow ids included with `0,0`). |
| `CellRef.TryFormat(col, row, span, out written)` / `CellRef.Format(col, row)` | A1-style id formatting from numeric addresses; the `TryFormat` span form allocates nothing (18 chars always suffice). |
| `GetValueReader(sheetName)` | Numeric-address bulk reader for one sheet → `SheetValueReader`; `GetValue(column, row)` serves memoized values with no per-cell id string, A1 parse or sheet-name hash — misses evaluate on demand, identical to `GetCellValue`. See [Bulk reads](#bulk-reads-getvaluereader). |
| `ComputeAll()` | Eagerly evaluates every cell (the "calculate now" counterpart to lazy `GetCellValue`), filling the cache so a later warm save carries computed values. Runs on a large stack for deep chains; a second call is all hits. After edits, call `InvalidateCache()` first. |
| `InvalidateCache()` | Explicitly flushes the whole memoization cache (required after edits); also resets the volatile epoch. |
| `Recalculate()` | Refreshes only volatile cells (see [Volatile functions](#volatile-functions)); keeps every stable cell cached. |
| `TimeProvider` | Injectable clock for `NOW`/`TODAY` (defaults to `TimeProvider.System`, read in local time). |
| `RandomSeed` | Optional `int?` seed for `RAND`/`RANDBETWEEN` (fixed value → reproducible runs). |
| `RegisterFunction(name, fn)` / `TryGetFunction(name, out fn)` | Custom-function registry ([guide](custom-functions.md)). |
| `Save(path)` / `SaveAsync(path)` / `Load(path)` / `LoadAsync(path)` | MemoryPack serialization ([guide](serialization.md)). |
| `RunWithLargeStack(work)` (static) | Runs an evaluation batch on a large-stack thread ([guide](performance.md)). |

## Sheet

A `Sheet` maps cell ids (`"A1"`, `"B12"`, …) to `Expression` nodes:

```csharp
using Danfma.MySheet.Expressions;
using Danfma.MySheet.Parsing;

sheet["A1"] = new NumberValue(1);                       // set: stores the expression
sheet["B1"] = ExpressionParser.Parse("=A1*2", sheet);

Expression cell = sheet["A1"];      // get: never throws — a missing cell reads as BlankValue.Instance
bool exists = sheet.ContainsKey("C1");                  // false
bool found = sheet.TryGetValue("A1", out var stored);   // true
bool removed = sheet.Remove("B1");                       // delete a cell → true if it existed

foreach (var (id, expression) in sheet) { /* iterate stored cells */ }
```

- The **getter never throws**: reading an id that was never set returns `BlankValue.Instance`, which
  evaluates to a blank — exactly how Excel treats an empty cell.
- **Writing goes through the indexer `set`, and deleting through `Remove`** — the two, and only, paths
  that change a sheet's cells. `Remove(id)` returns `true` when a cell was there and `false` for a no-op.
  Like the `set`, `Remove` does not clear the memoization cache on its own: after editing (writing or
  removing), call `workbook.InvalidateCache()` before reading again for the change to be observed.
- **`Cells` is a read-only view** (`IReadOnlyDictionary<string, Expression>`) of the stored cells —
  enumerable and indexable for reading (`sheet.Cells["A1"]`, `sheet.Cells.Count`), but not mutable;
  mutate through the indexer and `Remove`. (Callers upgrading from 2.x that mutated `Cells` directly:
  see [Migrating to 3.0](migrating-to-3.0.md).)
- `Keys`, `Values` and `Count` expose only the cells that were actually stored.
- Cell ids are plain A1-style strings. The parser normalizes them to upper-case and strips absolute
  markers (`$A$1` → `A1`); when setting cells through the indexer yourself, use the normalized form
  (`"A1"`, not `"a1"`).

## Expressions

Every cell holds an `Expression` — an immutable record. Literals, references and operators live in
`Danfma.MySheet.Expressions`; the built-in function nodes live in per-category child namespaces
(`Danfma.MySheet.Expressions.Mathematics`, `.Logical`, `.Statistical`, `.Text`, `.Information`,
`.Lookup`, `.Financial` — see [Migrating to 2.0](migrating-to-2.0.md)). Together they form a tree.

### Parsing

`ExpressionParser.Parse(text, sheet)` converts a cell entry into an expression, using the sheet as the
context for unqualified references:

```csharp
using Danfma.MySheet.Parsing;

var formula = ExpressionParser.Parse("=SUM(A1:A10) * 1.1", sheet);   // expression tree
var number = ExpressionParser.Parse("42.5", sheet);                  // NumberValue
var flag = ExpressionParser.Parse("true", sheet);                    // BooleanValue
var text = ExpressionParser.Parse("hello", sheet);                   // StringValue
var blank = ExpressionParser.Parse("", sheet);                       // BlankValue
```

Rules:

- Entries starting with `=` are parsed as formulas (a Pratt / top-down operator-precedence parser).
- Anything else is a literal: number if it parses as one (invariant culture), then boolean
  (`true`/`false`), otherwise text.
- **Syntax errors throw `ParseException`.** The exception is structured: `Kind` (a `ParseErrorKind` —
  `UnexpectedCharacter`, `UnterminatedString`, `UnterminatedQuotedName`, `UnexpectedToken`, `ExpectedToken`,
  `ExpectedCellReference`, `InvalidArgumentCount`, `NestingTooDeep`; `Unspecified` only when the exception is
  built through the legacy `(message, position)` constructor), `Token` (the offending token's text,
  empty when the parser ran out of input) and `Position` (0-based offset into the formula **body**, i.e. the
  text after the leading `=`, so `=1 2` reports position 2 for the `2`). Catching `ParseException` is by
  itself the "syntax vs. semantic" distinction — see the next bullet. Built-in functions also validate their
  argument count at parse time — `=ROUND(1)` throws, just as Excel would reject it at entry.
- **Semantic errors do not throw** — an unknown function evaluates to `#NAME?`, a bad reference to
  `#REF!`, and so on, as `ComputedValue` errors. A **reference to a sheet that does not exist**
  (`=Ghost!A1`, `SUM(Ghost!A:A)`) is one such bad reference: it resolves to `#REF!` — never a thrown
  `KeyNotFoundException` — so one dangling cross-sheet reference cannot abort a whole-workbook batch. A
  missing sheet is a *structural* error, so it propagates through **every** consuming function — the
  aggregations, the error-ignoring `COUNT` family (`COUNT`/`COUNTA`/`COUNTIF` over a ghost sheet are
  `#REF!`, not `0`), the lookups (`VLOOKUP`/`MATCH`/`XLOOKUP`/`INDEX`), `SUMPRODUCT` and the statistical
  pairs, the financial cash-flow functions (`NPV`/`IRR`/`XIRR`/`MIRR`), the text joins (`CONCAT`/`TEXTJOIN`)
  and the rest. A value error *inside* a cell of an existing sheet keeps its usual per-function policy
  (`COUNT` ignores it, `SUM` propagates it), and an empty range over an *existing* sheet is still a value
  outcome (`MATCH` over it is `#N/A`, not `#REF!`).

### Building trees in code

You can construct expressions directly — useful for programmatic workbooks and tests:

```csharp
using Danfma.MySheet.Expressions;
using static Danfma.MySheet.Expressions.Expression;

sheet["A1"] = Number(10);
sheet["A2"] = Number(20);
sheet["A3"] = Sum(Cell("A1", sheet), Cell("A2", sheet));
sheet["A4"] = Add(Cell("A3", sheet), Number(5));
sheet["A5"] = Sum(Range("A1", "A2", sheet));
```

The `Expression` base class provides factory helpers (`Number`, `String`, `Cell`, `Range`, `Sum`,
`Average`, `Min`, `Max`, `Count`, `Add`, `Subtract`, `Divide`, `Power`, `GreaterThan`, `Negate`,
`Plus`), and every node type is a public record you can `new` up directly (`new NumberValue(1)`,
`new BinaryOperation(BinaryOperator.Multiply, left, right)`, …). To `new` up a function record
directly, import its category namespace (e.g. `using Danfma.MySheet.Expressions.Mathematics;` for
`new Sum(…)`).

### Evaluating

`Evaluate` is the single evaluation contract. It returns a [`ComputedValue`](computed-value.md), with no
boxing for numeric results:

```csharp
ComputedValue direct = sheet["A3"].Evaluate(workbook);          // evaluates the tree
ComputedValue cached = workbook.GetCellValue("Sheet1", "A3");   // memoized per cell
```

Prefer `GetCellValue` when reading cells: it caches the result, and any `CellReference` inside a formula
goes through the same cache, so shared cells are computed once. `Evaluate` on an expression instance is
the right tool for ad-hoc expressions that are not stored in a cell:

```csharp
var adHoc = ExpressionParser.Parse("=AVERAGE(A1:A2) > 10", sheet);
bool isHigh = adHoc.Evaluate(workbook).ToBoolean();
```

There is no other evaluation API: callers that need a loosely-typed `object?` call `.AsObject()` on the
result.

### Bulk reads: `GetValueReader`

An extraction loop that builds ids — `GetCellValue(sheetName, "C" + r)` — pays three per-cell costs the
result does not need: the id-string allocation, the A1 parse and a sheet-name hash lookup.
`Workbook.GetValueReader(sheetName)` resolves the sheet handle once and reads by 1-based numeric address:

```csharp
var reader = workbook.GetValueReader("Results");

for (var row = 2; row <= 60_001; row++)
{
    if (reader.GetValue(column: 2, row).TryGetNumber(out var value))
    {
        total += value;
    }
}
```

Semantics are identical to `GetCellValue`: a hit is a direct read of the paged value store; a miss
evaluates on demand (memoization, cycle guard), so literals and never-computed formulas are served too.
`InvalidateCache()` applies the same way, and the reader instance stays valid across invalidations.
Measured on a 360k-cell extraction: `29.8 ms / 24.2 MB` allocated with per-cell ids → `6.9 ms / 0 bytes`
with the reader.

To drive the loop from the sheet itself (instead of known bounds), enumerate the populated cells. The
fully allocation-free pipeline — including rendering the id text, e.g. for a JSON field — combines
`CellAddresses`, `CellRef.TryFormat` and the reader:

```csharp
var reader = workbook.GetValueReader("Results");
Span<char> id = stackalloc char[18];

foreach (var (column, row) in sheet.CellAddresses)      // struct enumerator, zero alloc
{
    CellRef.TryFormat(column, row, id, out var length); // id text without a string
    jsonWriter.WriteString(id[..length], reader.GetValue(column, row).ToDouble());
}
```

When you want the id as a `string` anyway, `sheet.EnumerateCells()` yields `(Id, Column, Row)` — the
library derives the canonical id once per cell (one string each), and the numeric address pairs
directly with the reader. Enumeration order is insertion order, not row-major; sort if you need
deterministic output.

### Formula results are never blank (Excel parity)

At the **cell boundary** — `GetCellValue` — a formula result is never blank, exactly like Excel: when a
cell that HAS content (its expression is not the empty `BlankValue`) evaluates to blank, `GetCellValue`
returns `ComputedValue.Number(0)`, and that coerced `0` is what enters the cache. A cell that is truly
empty (its expression is `BlankValue.Instance`, e.g. an id that was never set) stays blank.

```csharp
sheet["A1"] = Cell("F10", sheet);                       // =F10, F10 empty
workbook.GetCellValue("Sheet1", "A1").ToDouble();        // 0   (formula result coerced)
workbook.GetCellValue("Sheet1", "F10").Kind;             // Blank (truly empty cell)

sheet["A2"] = ExpressionParser.Parse("=IF(TRUE, F10)", sheet);
workbook.GetCellValue("Sheet1", "A2").ToDouble();        // 0   (blank branch coerced)
```

The coercion belongs to the **cell**, not the expression: `Evaluate` keeps blank INTERNALLY, so blank
still compares as `""`/`0`/`FALSE` inside an expression. This preserves the internal semantics while
matching Excel at the display boundary:

```csharp
sheet["A3"] = ExpressionParser.Parse("=IF(F10=\"\",1,2)", sheet);
workbook.GetCellValue("Sheet1", "A3").ToDouble();        // 1   (F10 empty still equals "" internally)
sheet["A4"] = ExpressionParser.Parse("=F10&\"\"", sheet);
workbook.GetCellValue("Sheet1", "A4").ToText();          // ""  (result is text, not blank → not coerced)
```

The parity effects cascade, all matching Excel: `ISBLANK(A1)` with `A1 = "=F10"` is **FALSE** (A1 is now
0), `COUNT` counts a formula-empty cell (0 is a number) while `COUNTBLANK` no longer does, and the
`SaveAsExcel` `ValuesOnly` export writes `0` for a formula-empty cell instead of omitting it.

## Operators

MySheet parses the Excel operator set. Binding powers (precedence) from loosest to tightest:

| Precedence | Operators | Notes |
| --- | --- | --- |
| 1 (loosest) | `=` `<>` `<` `>` `<=` `>=` | Comparisons, with Excel's cross-type ordering (numbers < text < logicals). |
| 2 | `&` | Text concatenation. |
| 3 | `+` `-` | Addition, subtraction. |
| 4 | `*` `/` | Multiplication, division. |
| 5 | `^` | Exponentiation (parsed right-associatively). |
| 6 | `%` | Postfix percent: `50%` is `0.5`. |
| 7 | unary `-` `+` | Unary prefix binds tighter than `^`, so `-2^2` is `(-2)^2 = 4`, matching Excel. Unary `-` (and postfix `%`) coerce to a number (`-"abc"` is `#VALUE!`); unary `+` is Excel's legacy no-op and returns the operand **unchanged, type included** — `=+A1` on a text cell is that text, `+TRUE` stays a boolean, a reference stays a reference (`SUM(+A1:A3)`), only a blank becomes `0`. |
| 8 (tightest) | `:` | Range construction. |

Plus grouping with `( )`. Division by zero yields `#DIV/0!`; type mismatches yield `#VALUE!`.

## References

```csharp
ExpressionParser.Parse("=A1", sheet);                 // same-sheet cell
ExpressionParser.Parse("=$A$1+A2", sheet);            // absolute markers accepted (and normalized away)
ExpressionParser.Parse("=Sheet2!A1", sheet);          // sheet-qualified
ExpressionParser.Parse("='My Sheet'!A1:B2", sheet);   // quoted sheet name, range
ExpressionParser.Parse("=SUM((A1:A3, C1:C3))", sheet); // reference union (inside parentheses)
```

- Unqualified references resolve against the sheet passed to `Parse`.
- A range (`A1:B2`) requires cell references on both sides and lives on the start cell's sheet
  (`Sheet2!A1:B2` is entirely on `Sheet2`).
- `$` markers identify the same cell — MySheet does not do copy/fill, so absolute vs. relative has no
  behavioral effect and the marker is not preserved.
- A bare range is consumed by the functions that accept it (`SUM`, `COUNT`, lookups, …); it has no scalar
  value of its own. Evaluated directly (`Parse("=A1:B2", sheet).Evaluate(workbook)`) it is `#VALUE!`, but
  the same formula stored in a **cell** is
  [implicitly intersected](#implicit-intersection-at-the-cell-boundary) with that cell's row and column.
- A bare name that is not a cell id (e.g. `=total`) is a `NameReference` — it resolves against `LET`
  bindings and workbook [named ranges](#named-ranges) at evaluation time, and yields `#NAME?` if unbound.

## Whole-column and whole-row references

MySheet supports references that are **open** (unbounded) on at least one side: a whole column, a whole
row, and the one-sided mixed forms.

```csharp
ExpressionParser.Parse("=SUM(A:A)", sheet);     // whole column A
ExpressionParser.Parse("=SUM(A:C)", sheet);     // columns A..C
ExpressionParser.Parse("=SUM(1:1)", sheet);     // whole row 1
ExpressionParser.Parse("=SUM(1:5)", sheet);     // rows 1..5
ExpressionParser.Parse("=SUM($1:$1)", sheet);   // '$' on a row endpoint is a fill marker only: same as 1:1
ExpressionParser.Parse("=SUM(A2:A)", sheet);    // column A from row 2 downward
ExpressionParser.Parse("=SUM(A:A10)", sheet);   // column A up to row 10
ExpressionParser.Parse("=SUM(A1:C)", sheet);    // columns A..C from row 1 downward
ExpressionParser.Parse("=SUM(Data!A:A)", main); // sheet-qualified; $A:$A is accepted and normalized
```

These parse to a single `OpenRangeReference(int? ColMin, int? ColMax, int? RowMin, int? RowMax, string
SheetName)`; each limit is `null` where that side is open. The **left** endpoint gives the lower bounds
(`ColMin`/`RowMin`), the **right** the upper bounds (`ColMax`/`RowMax`); an endpoint that names only a
column informs no row (and vice-versa), so that axis stays open on that side. When all four limits are
known the parser produces a plain `RangeReference` instead.

**Semantics — populated cells.** An open range means *the populated cells within the limits*, not a fixed
grid. Enumeration scans the sheet's stored cells and keeps those whose (column, row) fall inside the
non-null bounds; blank cells contribute `0`, so `SUM(A:A)` matches Excel while never materializing the
1,048,576-row column. An empty column sums to `0`. This keeps whole-column aggregation cheap on the sparse
model MySheet uses.

**`:` forces reference semantics.** A letters-only endpoint adjacent to `:` is a **column**, and an
integer endpoint a **row**, even when a defined name of the same spelling exists — so `Sales:Sales` is the
column `SALES`, not the named range. (This edge only matters if you name something after a column label.)

**`ROWS` / `COLUMNS` — a documented divergence from Excel.** Excel reports the fixed grid size
(`ROWS(A:A)` = 1,048,576). A gridless model has no such grid, so MySheet uses the **populated extent** on
an open axis and the **exact structural count** on a bounded axis:

| Formula        | MySheet result                              | Excel      |
| -------------- | ------------------------------------------- | ---------- |
| `ROWS(A:A)`    | max populated row − min populated row + 1 (0 if empty) | 1,048,576 |
| `COLUMNS(A:C)` | `3` (structural, exact)                     | `3`        |
| `COLUMNS(A:A)` | `1`                                         | `1`        |
| `ROWS(1:5)`    | `5` (structural)                            | `5`        |

**Reference consumers.** Where a concrete range is required — `VLOOKUP`/`HLOOKUP` (table), `INDEX`,
`OFFSET` (base) — an open range resolves to the **populated bounding box** within its limits; `AREAS`
counts it as one area and `ISREF` reports `true`. So `VLOOKUP(2, A:B, 2)` and `INDEX(A:A, 3)` work.

**Out of scope.** Spatial intersection of two open ranges is not modeled.

## Implicit intersection at the cell boundary

A formula whose **final** value is a multi-cell reference is not an error in a cell. MySheet applies
Excel's implicit intersection — the `@` operator — against the formula cell's own row and column:

| The cell's formula denotes | The cell shows |
| --- | --- |
| a **single-column** range spanning the formula's **row** | that row's cell in the column — `=A1:A3` in `C3` is `A3` |
| a **single-row** range spanning the formula's **column** | that column's cell in the row — `=A1:C1` in `B5` is `B1` |
| a **1x1** range | that cell, wherever the formula sits — `=A1:A1` in `Z99` is `A1` |
| a range the formula's row/column falls **outside** of | `#VALUE!` — `=A1:A3` in `C9` |
| a range wider than one cell on **both** axes | `#VALUE!` — `=A1:C3` in `B2` |
| a **union** of areas | `#VALUE!` — `=(A1:A3,B1:B3)` has no single row/column axis to intersect |

Details:

- **Declared bounds, not populated ones.** An [open range](#whole-column-and-whole-row-references) uses the
  bounds it declares, so `=A:A` in row 7 is `A7` even when `A7` is empty (the blank then becomes `0` by the
  [never-blank rule](#formula-results-are-never-blank-excel-parity)), `=1:1` in `B7` is `B1`, and `=A2:A` is
  `#VALUE!` in row 1 but `A3` in row 3. This is deliberately *not* the populated extent that `ROWS`/
  `COLUMNS` use.
- **Positional and sheet-independent.** Only the formula cell's row and column number enter the rule:
  `=Sheet1!A1:A3` typed in `Sheet2!C2` is `Sheet1!A2`. *(Inference, not a measurement: Excel defines `@`
  purely in terms of row and column with no sheet term; this cross-sheet case was not checked against
  Excel.)*
- **Everything that denotes a reference follows the same table**, not only a literal range — `=MyName`,
  `=INDIRECT("MyName")`, `=OFFSET(A1,0,0,3,1)`, `=CHOOSE(1,A1:A3)`, `=+A1:A3` and `=LET(x,A1:A3,x)` all
  intersect. Before this rule they stored a reference-kind value that every typed accessor read back as
  blank.
- **Inside a formula nothing changes.** `=SUM(A1:A3)` is still a sum over three cells: the *consumer*, not
  the cell, decides what a multi-cell reference means. Only a reference that survives as the cell's final
  value is intersected.
- **The direct `Expression.Evaluate` path still yields `#VALUE!`.**
  `ExpressionParser.Parse("=A1:A3", sheet).Evaluate(workbook)` has no formula cell to intersect against.
  The rule lives in `Workbook.EvaluateCell`, which is the single choke point of every cell read
  (`GetCellValue`, the [value reader](#bulk-reads-getvaluereader), the warm-start snapshot, the `.xlsx`
  export) — so a `ComputedValueKind.Reference` can never be a cell's value.
- **A cell that intersects itself is `#REF!`.** `=A1:A3` in `A2` dereferences `A2`, the cell already on the
  evaluation stack, so the cycle guard answers `#REF!` (Excel raises a circular-reference dialog instead).
- **No spill.** The intersected cell is the whole result — MySheet never writes into neighbouring cells.

**Deviations worth knowing.** A 2-D range answers `#VALUE!` even when the formula cell sits inside the
rectangle. That follows the `@` rule as documented, but it is the likeliest point of divergence from a real
Excel build and was **not** measured; the alternative reading — the formula cell's own (column, row) when
the rectangle contains it — is one extra branch in `ImplicitIntersection`. A union is likewise deliberately
left at `#VALUE!`. A **computed array** never reaches this rule at all — a bare `=A1:C3*E1:E3` or `=LEN(A1:A3)`
in a cell is the operator's or the function's own `#VALUE!`, not an intersection — and Excel's answer for the
typed form of those is measured and recorded under
[implicit array arguments](#implicit-array-arguments); reconciling the two is the array half of this rule.

## Implicit array arguments

A few functions evaluate an **array-valued argument element-by-element**, reproducing Excel's implicit
(CSE) semantics — no `Ctrl+Shift+Enter`, no spilling, and no public array value: the vector lives only
inside the consuming function's evaluation. This closes the common `SUM(IF(range=…))` /
`SMALL(IF(range=…, ROW(range)))` idioms.

```csharp
ExpressionParser.Parse("=SUM(IF(B2:B5=\"Show\",1,0))", sheet);            // → 2 (count of matches)
ExpressionParser.Parse("=SMALL(IF(B2:B5=\"Show\",ROW(B2:B5)),1)", sheet); // → 3 (first matching row)
ExpressionParser.Parse("=INDEX(ROW(B2:B5),1)", sheet);                    // → 2 (row vector, indexed)
ExpressionParser.Parse("=INDEX(ROW($A:$A),4)", sheet);                    // → 4 (identity: nth row)
```

**Which entry mode the Excel numbers on this page come from.** Excel answers these shapes *differently*
depending on how the formula was entered. Typed normally it applies legacy implicit intersection **inside**
the argument; array-entered (`Ctrl+Shift+Enter`) it evaluates the whole array. So `SUM(A1:C3*E1:E3)` over the
fixture below is `#VALUE!` typed and **108** array-entered, and `SUM(ROW(A1:C3))` is **1** typed — `ROW` of a
rectangle answers its single top row there — against **6** array-entered. MySheet has no array-entry concept
at all: a formula is a formula, and every consumer described here implements the **array-entered** rule. Every
Excel figure quoted in this section and in its divergence list is therefore the array-entered one unless the
line says *typed* (all measured on Aspose.Cells 26.6.0, 2026-09-10). If you type one of these formulas into a
real Excel and compare, expect the typed answer rather than ours; the one place where MySheet's own answer
follows neither is [the cell boundary](#implicit-intersection-at-the-cell-boundary), below.

**Supported.** The consumers are the numeric aggregators (`SUM`, `COUNT`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`
and — through the same fold — `MEDIAN`, the `STDEV`/`VAR` family, `SMALL`, `LARGE`, the percentiles and
quartiles), `INDEX`, `ROWS`/`COLUMNS` (which report the array's *extent* rather than its values),
`COUNTA`/`CONCAT`/`TEXTJOIN` (which stream its elements), `SUMPRODUCT`, and the **array form** of
[`AGGREGATE`](function-reference.md) (`function_num` 14-19), where option 6 drops the `#DIV/0!` elements that
make a plain `SMALL` over the same vector fail. `SUBTOTAL` and AGGREGATE's *reference* form (1-13) are
deliberately **not** consumers: their arguments are `ref`s, and Excel rejects a computed array in one —
`SUBTOTAL(9,ROW(A1:A3))` and `AGGREGATE(9,4,ROW(A1:A3))` are `#VALUE!` there, measured on Aspose.Cells
26.6.0, which is exactly why AGGREGATE documents a second syntax for arrays. An argument is evaluated as an
array when it is a **closed-range** comparison (`B2:B5="Show"`), an `IF` whose condition is such an array
(with or without an else branch), `ROW`/`COLUMN` over a rectangle, one of the two **lifted** shapes
described further down, or one of the four [dynamic-array producers](#dynamic-array-producers)
(`FILTER`/`SORT`/`UNIQUE`/`SEQUENCE`). That rectangle may be written literally — `ROW(A1:C3)` is the 3x1 column `[1,2,3]`
and `COLUMN(A1:C3)` the 1x3 row `[1,2,3]`, one number per rank rather than one per cell, so
`SUM(ROW(A1:C3))` and `SUM(COLUMN(A1:C3))` are both 6 and `COUNT` of either is 3 (array-entered; typed, both
sums are 1) — or merely *denoted* by the argument — a
[defined name](#named-ranges) (`SUM(ROW(MyName))` = 6 and `COUNT(ROW(MyName))` = 3 for a name over three
rows, while `COUNT(MyName)` counts the cells' own values) or a `:` range with reference-returning endpoints
(`SUM(ROW(INDEX(A1:A3,1,1):A3))` = 6). A branch-less `IF` yields a
logical `FALSE` where the condition is false, and the aggregators ignore logicals/text (exactly why
`SMALL(IF(…))` skips the non-matching rows). The first per-element error wins, as in Excel.

**How two arrays of different shape combine (broadcasting).** A scalar broadcasts to every position — and so
does a **1x1 range**, which is a vector too: `SUM(A1:C3*E1:E1)` = 45. Otherwise each axis is decided on its
own, and an extent of 1 defers to the other side: an Nx1 **column** repeats across every column
(`SUM(A1:C3*E1:E3)` = 108), a 1xM **row** repeats down every row (`SUM(A1:C3*E5:G5)` = 960), and an Nx1
against a 1xM is the NxM **outer product** (`SUM(E1:E3*E5:G5)` = 360). `IF` folds its condition and *both*
branches into one extent by the same rule (`SUM(IF(E1:E3>1,A1:C3,0))` = 39, `SUM(IF(E1:E3>1,E5:G5,0))` = 120).
Where both extents exceed 1 and differ, the result takes the **larger** one and every position the shorter
operand does not cover is **`#N/A`** — a per-element error, not a whole-expression one, so a consumer that
skips or catches errors still answers: `SUM(A1:C3*H1:H2)` is `#N/A` while `COUNT(A1:C3*H1:H2)` = 6 and
`AGGREGATE(15,6,A1:C3*H1:H2,6)` = 12. `SUMPRODUCT` keeps its own stricter rule, and only for its **own**
argument list: those must match exactly, a 1x1 included — `SUMPRODUCT(A1:C3,E1:E3)` and
`SUMPRODUCT(A1:C3,E1:E1)` are both `#VALUE!`, where the operator would have broadcast both — while a broadcast
written *inside* one argument is consumed as the array it computes (`SUMPRODUCT(A1:C3*E1:E3)` = 108). Fixture:
`A1:C3` = 1…9 row-major, `E1:E3` = 1,2,3, `E5:G5` = 10,20,30, `H1:H2` = 1,2. Every figure here is measured on
Aspose.Cells 26.6.0, 2026-09-10, array-entered — typed, every `SUM` form above is `#VALUE!` there and
`COUNT(A1:C3*H1:H2)` is 0, while the `AGGREGATE` and `SUMPRODUCT` forms answer the same in both modes
(they are array-native under either entry) — and pinned by `VectorBroadcastingTests`,
`MiniCseConsumerTests` and `MathAggregateTests`.

**Lifted unary operators and scalar functions.** Two further shapes become arrays wherever one of the
consumers above asks for one, by applying a scalar body element by element:

- a unary `-` or `%` over an array — over `A1:A3` = 1,2,3: `SUM(-A1:A3)` = -6, `SUM(A1:A3%)` = 0.06,
  `SUM(-(A1:A3>1))` = -2, and the double-negation idiom `SUMPRODUCT(--(A1:A3>1))` = 2;
- any **pure-scalar built-in** with at least one array argument — `SUM(LEN(D7:F9))` = 7 over a 3x3 rectangle
  holding `"abc"`, `"def"` and one space, `COUNT(LEN(D7:F9))` = 9 (nine lengths, the blanks included),
  nested lifts (`SUM(LEN(TRIM(D7:F9)))` = 6), `SUM(ABS(A1:A3*-1))` = 6, `SUM(ROUND(A1:A3,0))` = 6,
  `SUM(ISNUMBER(A1:A3)*1)` = 3, `SUM(IFERROR(A1:A3,0))` = 6, and the whole worksheet idiom
  `IF(SUMPRODUCT(--(LEN(TRIM($D$7:$F$9))>0))>0,"Show","Hide")`.

**180 of the 310 registered built-ins** are liftable: the pure-scalar ones (text, mathematics, financial,
date, information, the scalar statistics helpers (`FISHER`, `PERMUT`, `PHI`, `STANDARDIZE`, …),
`IFERROR`/`IFNA`/`IFS`/`NOT`/`SWITCH`, `ADDRESS`). The
other 130 are **range-aware** and MySheet never lifts them, because they consume ranges or arrays themselves —
or, for the four [dynamic-array producers](#dynamic-array-producers), because they *produce* one —
`SUM`, `COUNT`, `INDEX`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `AREAS`, `SUMPRODUCT`, `SUBTOTAL`, `AGGREGATE`,
`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`,
`VLOOKUP`, `MATCH`, `OFFSET`, `INDIRECT`, `IF`, `LET`, `RANDBETWEEN`, `AND`/`OR`/`XOR`, the cash-flow series
(`NPV`, `IRR`, …), the whole-population and paired-array statistics (`RANK`, `MODE`, `CORREL`, `SUMXMY2`, …)
and the criteria family. That is MySheet's rule and **not** Excel's: Excel lifts a range-aware function too,
over the slots that take a *scalar*, while still consuming the range in the slot that takes one — the last of
the known divergences below, with the twelve measured cases. A
[custom function](custom-functions.md) is never lifted either — it has no registry entry, so it
stays a scalar evaluated once.

Inside a lifted call:

- **Scalar arguments broadcast**, and each is evaluated exactly **once** per evaluation rather than once per
  element: `SUM(ROUND(A1:A3,0))` reads the `0` once, and a volatile or a host function in a scalar slot is
  called a single time for the whole vector.
- **Two array arguments broadcast per axis**, by the same rule the operator's operands follow (the
  broadcasting paragraph above). Equal shapes pair position by position (`SUM(ROUND(A1:A3,B1:B3))` = 6 over
  1,2,3 and 10,20,30); an extent of 1 defers to the other side, so `SUM(ROUND(A1:C3,E1:E3))` — 3x3 against
  3x1 — is 45; and where both extents exceed 1 and differ, the uncovered positions are `#N/A`, so
  `SUM(ROUND(A1:C3,H1:H2))` is `#N/A` while `COUNT(ROUND(A1:C3,H1:H2))` = 6. An earlier release handed the
  body a `#VALUE!` marker for *every* element instead; that is gone, which changes what an
  error-*consuming* body sees — `SUM(LEN(IFERROR(LEFT(D7:F9,E6:E8),"zz")))` is 0, not 18, because the
  broadcast blank column makes every element `LEFT(x,0)` = `""` and leaves `IFERROR` nothing to recover
  (all array-entered, measured on Aspose.Cells 26.6.0).
- **An omitted argument keeps the function's own default.** `FIXED(A1:A3,,TRUE)` still formats two decimals
  per element: an omitted slot stays a blank literal in the tree rather than becoming a per-element slot, so
  the function's "argument not supplied" branch still fires.
- **Errors propagate per element**, the first in row-major scan order winning: `SUM(ABS(1/(A1:A3-2)))` is
  `#DIV/0!` and `SUM(LEN(Ghost!A1:A3))` is `#REF!` — the lifted path reads cells, so the missing-sheet rule
  applies (unlike `SUM(ROW(Ghost!A1:A3))`, the divergence below). Text where a number is required makes that
  element `#VALUE!` (`SUM(ABS(B1:B3))` with `B2` = `"x"`), and a blank element coerces to `0`, so
  `SUM(LEN(A1:A3))` over three empty cells is `0` while `COUNT(LEN(A1:A3))` is `3`.

### Dynamic array producers

`FILTER`, `SORT`, `UNIQUE` and `SEQUENCE` are the other direction of the same machinery: instead of turning a
range into an array, they **produce** one. Anywhere a consumer above accepts an array, one of these four can
stand in its place, and they compose with each other and with every shape described above.

```csharp
// A1:A3 = 5, 0, 9;  B1:B3 = 1, 2, 3;  Q1:Q4 = 9, 5, 9, 0
ExpressionParser.Parse("=SUM(FILTER(A1:A3,A1:A3>0))", sheet);   // → 14  (0 fails the predicate)
ExpressionParser.Parse("=ROWS(FILTER(A1:A3,A1:A3>0))", sheet);  // → 2   (how many rows matched)
ExpressionParser.Parse("=INDEX(SORT(A1:A3,1,-1),1)", sheet);    // → 9   (the largest)
ExpressionParser.Parse("=COUNTA(UNIQUE(Q1:Q4))", sheet);        // → 3   (distinct count)
ExpressionParser.Parse("=SUM(SEQUENCE(5))", sheet);             // → 15  (1+2+3+4+5)
ExpressionParser.Parse("=INDEX(SEQUENCE(2,3),2,2)", sheet);     // → 5   (row-major fill)
```

Each function's own arguments, defaults, coercions and error rules are in the function reference —
`FILTER`/`SORT`/`UNIQUE` under [Lookup and reference](function-reference.md#lookup-and-reference-20),
`SEQUENCE` under [Math and trigonometry](function-reference.md#math-and-trigonometry-76). What this section
covers is how they behave *as arrays*.

**Every consumer reads them, with no per-function special case.** They are operands in the same lazy tree, so
the whole **Supported** list above applies unchanged. Measured on this engine over `A1:A3` = 5, 0, 9:
`SUM(FILTER(A1:A3,A1:A3>0))` = 14, `COUNT` 2, `AVERAGE` 7, `MIN` 5, `MAX` 9, `PRODUCT` 45,
`SMALL(…,1)` 5, `LARGE(…,1)` 9, `MEDIAN(SEQUENCE(5))` = 3, `PERCENTILE(SEQUENCE(5),0.5)` = 3,
`INDEX(SORT(A1:A3),1)` = 0, `ROWS(FILTER(A1:A3,A1:A3>0))` = 2, `COLUMNS(SEQUENCE(2,3))` = 3,
`COUNTA(UNIQUE(Q1:Q4))` = 3, `CONCAT(SEQUENCE(3))` = `"123"`,
`TEXTJOIN(",",TRUE,FILTER(A1:A3,A1:A3>0))` = `"5,9"`, `SUMPRODUCT(SEQUENCE(3))` = 6 and
`AGGREGATE(15,6,FILTER(A1:A3,A1:A3>0),1)` = 5.

Two of those consumers are new arrivals, and they now answer for **every** computed array rather than only for
a producer: `ROWS`/`COLUMNS` report the extent of an operator's result or a lifted call too — `ROWS(A1:C3*2)`,
`ROWS(A1:C3*H1:H2)` (the *broadcast* extent), `ROWS(LEN(A1:A3))`, `ROWS(-A1:A3)` and `ROWS(ROW(A1:C3))` are all
3, where they used to be `#VALUE!` or `1` — and `COUNTA`/`CONCAT`/`TEXTJOIN` stream the elements of one
(`COUNTA(IF(A1:A3>0,A1:A3))` = 3, `CONCAT(A1:A3*2)` = `"10018"`). A producer's own 1x1 **error** is reported
as itself rather than as a shape of 1 (`ROWS(FILTER(A1:A3,A1:A3>100))` = `#CALC!`,
`ROWS(SEQUENCE(-1))` = `#VALUE!`), while an error *element* inside a rectangle does not hide the shape
(`ROWS(SORT(E1:E3))` = 3).

**Two members of the flattening family are exceptions, for two different reasons**, and it matters which:

- **`CONCATENATE` takes the top-left, it does not expand.** It joins scalars, so a producer argument reaches
  the producer's own cell answer: `CONCATENATE(SEQUENCE(3))` is `"1"` where `CONCAT(SEQUENCE(3))` is `"123"`.
  That is Excel's answer too, in both entry modes (measured on Aspose.Cells 26.6.0, 2026-09-10).
- **`COUNTBLANK` refuses a computed array outright**, with `#REF!`, before any streaming can happen:
  `COUNTBLANK(SEQUENCE(3))` is `#REF!` — the same rule and the same error the criteria family applies (the
  bullet under **Not supported** below), and the same answer Excel gives in both modes.

**They compose, and the composition needed no code of its own** — a producer under an operator, a lifted
function, an `IF`, or under another producer, is reached by the same recursive builder. All measured on this
engine and equal to Excel array-entered: `SUM(SORT(FILTER(A1:A3,A1:A3>0)))` = 14,
`SUM(FILTER(A1:A3,A1:A3>0)*2)` = 28, `SUM(LEN(FILTER(A1:A3,A1:A3>0)))` = 2,
`SUM(FILTER(SEQUENCE(5),SEQUENCE(5)>2))` = 12, `ROWS(UNIQUE(FILTER(A1:A3,A1:A3>0)))` = 2 and
`SUM(IF(TRUE,SEQUENCE(3),0))` = 6. Broadcasting applies as everywhere else, so a producer whose extent is
shorter than what it is combined with leaves `#N/A` in the positions it does not cover:
`SUM(FILTER(A1:A3,A1:A3>0)*B1:B3)` is `#N/A` — a 2x1 selection against a 3x1 range — while
`COUNT` of the same expression is 2.

**An empty result is `#CALC!`, never an empty array.** `SUM(FILTER(A1:A3,A1:A3>100))` and
`ROWS(FILTER(A1:A3,A1:A3>100))` are [`#CALC!`](computed-value.md), Excel's empty-array error, with
`ERROR.TYPE` **14**; supply `FILTER`'s third argument to answer something else
(`SUM(FILTER(A1:A3,A1:A3>100,0))` = 0). A 1x1 result — an empty one included — broadcasts like a scalar, so
`SUM(FILTER(A1:A3,A1:A3>100,7)*A1:A3)` is 98, the 7 against every cell. As always, `COUNT` and `COUNTA`
*discard* errors rather than propagating them, so `COUNT(FILTER(A1:A3,A1:A3>100))` is `0` — which is Excel's
answer there too.

**Blanks survive the selection as blanks; nothing is normalized to zero.** A blank source cell stays blank
through all three selectors, so `COUNTA(FILTER(A5:A8,A5:A8<>"zzz"))` = `COUNTA(A5:A8)` = 3 over
7, blank, `"t"`, 7, `ISBLANK` of a kept blank is `TRUE`, `UNIQUE` treats a blank as its **own** key (equal to
neither `0` nor `""` nor `FALSE`, so `ROWS(UNIQUE(A5:A8))` = 3) and `SORT` puts blanks **last** in both
directions. Every one of those matches Excel in both entry modes, measured 2026-09-10 — a RANGE argument and
a PRODUCER result agree, and there is no seam between them.

**Bare in a cell: the `@` rule — and MySheet does not spill.**

A cell holds one scalar value, so a producer written as a cell's whole formula shows the array's **top-left
element** and nothing is written to the cells around it. That is Excel's `@`-on-an-array rule (Microsoft,
"Implicit intersection operator: @": for an array Excel "picks the top-left value"), and it is exactly what
Excel itself writes when it upgrades a legacy formula to `=@FILTER(...)`. Measured on Aspose.Cells 26.6.0,
2026-09-10, one formula per workbook, in a cell whose row is *inside* the source range (`D2`) and in one
*outside* it (`D5`), plain and array-entered agreeing on every row — and MySheet answers the same in both
cells:

| In a cell | Excel and MySheet |
| --- | --- |
| `=SEQUENCE(5)` | `1` |
| `=SEQUENCE(2,3,7,1)` | `7` |
| `=FILTER(A1:A3,A1:A3>0)` | `5` |
| `=SORT(A1:A3)` | `0` (the sorted top-left) |
| `=UNIQUE(A1:A3)` | `5` |
| `=SEQUENCE(2,3)*10` | `10` |
| `=FILTER(A1:A3,A1:A3>100)` | `#CALC!` |

Unlike a bare range, the answer does **not** depend on where the formula sits: a producer has no worksheet
position to intersect with, which is what distinguishes the array half of the rule from
[the range half](#implicit-intersection-at-the-cell-boundary).

**So `=SEQUENCE(5)` shows `1` rather than filling five cells with 1..5, and `=FILTER(A:A,B:B>0)` shows one
value rather than a list.** That is the single most surprising consequence of this feature and it is worth
stating plainly rather than discovering: MySheet has **no spill model** — there is no way for one formula to
write into cells it does not own — so the top-left is the whole answer. If you want the list, consume it:
`ROWS(FILTER(...))` for the count, `INDEX(FILTER(...),k)` for the *k*th match,
`TEXTJOIN(",",TRUE,FILTER(...))` for all of them in one cell.

**The other deviations, in full.** Each is pinned by a test — with Excel's own number in the test wherever the
two engines differ — so closing one is always a deliberate edit.

- **No spill**, as above. It is the one deviation that is structural rather than a choice: a cell is one
  scalar. Note that the *value shown* agrees with Excel in both entry modes — what differs is that modern
  Excel would also fill the neighbouring cells.
- **An open or whole-column argument is refused.** The cost guard that refuses an open range in an array
  position applies to a producer's own arguments too, so `SUM(FILTER(A:A,A:A>0))` is `#VALUE!` here where
  Excel answers 14 (measured 2026-09-10, both entry modes, over a column A holding only `A1:A3` = 5, 0, 9).
  The refusal has a **silent half** worth knowing: `COUNT` discards the error channel, so
  `COUNT(FILTER(A:A,A:A>0))` is `0` here against Excel's 2 on that fixture — a wrong number rather than an
  error. Bound the range (`A1:A100000`) and it works. Making the open form work needs a shared-bounds rule,
  because `array` and `include` would otherwise be bounded independently and could disagree on row count.
- **`SEQUENCE` has a size cap that Excel does not.** `rows > 1048576`, `columns > 16384` or
  `rows * columns > 1048576` answers `#NUM!`; exactly at the cap is allowed. Excel has no cap in a consumed
  position — `ROWS(SEQUENCE(1048577))` is 1048577 and `COLUMNS(SEQUENCE(1,16385))` is 16385 there (measured
  2026-09-10, both modes) — but the element stream is lazy while every consumer walks every element, so
  without the cap `SUM(SEQUENCE(1000000,10000))` would hang instead of answering.
- **`UNIQUE(…, exactly_once)` follows Microsoft's page, not the measured oracle.** Over `Q1:Q4` = 9, 5, 9, 0
  the rows occurring exactly once are 5 and 0, and that is what MySheet returns (`ROWS` 2, `SUM` 5). The
  oracle keeps the *distinct*-count shape and pads it by repeating the last kept value — `ROWS` **3** with
  `SUM` **5**, i.e. rows 5, 0, 0 — so its `UNIQUE` result contains a duplicate, which its own row count
  contradicts. Where the oracle contradicts itself the documented rule wins; the measurement is recorded
  beside the test so the decision can be revisited.
- **`AVERAGE` over `UNIQUE` follows the page too, for the same reason.** Over that same `Q1:Q4` MySheet
  answers 14/3, which is its own `SUM` over its own `COUNT`. The oracle reports `SUM` **14**, `COUNT` **3**
  and `AVERAGE` **0** for the same expression — three answers that cannot all be right — and Microsoft's
  `AVERAGE` page is explicit that the mean is the sum over the count with zeros included.
- **A producer bound by `LET`, passed through `CHOOSE`, or through a unary `+`, collapses to its top-left.**
  `SUM(LET(x,FILTER(A1:A3,A1:A3>0),x))` is `5` here where Excel answers **14** in both entry modes (measured
  2026-09-10; `ROWS(LET(x,FILTER(…),x))` = 2, `SUM(CHOOSE(1,FILTER(…)))` and `SUM(+FILTER(…))` = 14 there).
  A binding is captured as a *value* before the element-wise evaluation can see it, so the producer's own
  cell answer is what gets bound. Use the producer directly in the consumer's argument slot. The loudest
  instance is a `LET`-bound producer in a criteria slot —
  `LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,">0"))` is `1` here — `COUNTIF` over the single collapsed element
  — against Excel's `#REF!` in both entry modes, with `SUM(f)` 5 against 14 and `ROWS(f)` 1 against 2 in the
  same shape. That row is this engine's one deliberately failing pin, red so that the fix turns it green
  rather than being discovered by accident.

**Which factory a new built-in uses (contributors).** The classification is one explicit flag per entry in
[`FunctionRegistry`](../Danfma.MySheet/Parsing/FunctionRegistry.cs): `Entry<T>(…)` registers a function that
consumes ranges/arrays itself and is never lifted, `Elementwise<T>(…)` a pure-scalar one the mini-CSE may
lift. **The default is `Entry<T>` — deny** — because the two mistakes are not symmetric: writing `Entry<T>`
where `Elementwise<T>` belonged only loses the optimization, while writing `Elementwise<T>` where `Entry<T>`
belonged makes the function answer from a single element of the rectangle it was meant to consume whole and
be **silently wrong**, with no error for anyone to notice. Since `Entry<T>` is also what an entry gets when
nobody chooses, *forgetting* the flag lands on the safe side by construction — the dangerous mistake is the
deliberate one.

The guard tests are precise about which of those two mistakes each one catches:

- A **forgotten** flag is safe by *construction*, not by a test. `Consumes` is the enum's zero value, so a
  new built-in registered through the default `Entry<T>` factory is never lifted, whatever it does with its
  arguments — the mistake costs nothing but the optimization. A test still requires such an entry to be
  *watched*: every range-aware entry must either be visible to the probe below or be named by hand, and one
  that is neither fails the suite carrying its own name.
- A **wrong** flag — `Elementwise<T>` on a range-aware built-in, the mistake that ships a silent wrong number
  — is caught by name. The exact set of the 180 `Elementwise` names is committed as a sorted roster, so
  adding a name fails the suite naming the newcomer and removing one fails naming the loss. A count would not
  do: a count survives a compensating swap, and it survives the honest-looking edit of flipping the factory
  and bumping the number. Measured, that edit used to leave the whole suite green.
- The same wrong flag is *also* caught with a diagnostic wherever the probe can see it, and that probe
  re-runs the derivation on every build. It sweeps every argument position of every arity from `MinArgs` to
  `MinArgs+3`, filling the remaining slots with a number, a text, a logical and a three-cell range in turn,
  and hands the entry three rectangles differing in position, shape and contents. A pure-scalar body answers
  identically for all three; a range-aware one does not, and the failure names the discriminating call. The
  sweep is still blind to **22** of the 130 range-aware entries — the ones that answer the same thing for
  every rectangle: the shape and reference tests (`AREAS`, `ISREF`, `ISFORMULA`, `FORMULATEXT`, `SHEET`,
  `TYPE`), `OFFSET`/`INDIRECT`, the design exclusions (`IF`, `LET`, `RANDBETWEEN`), and the folds that error
  identically on all three (`AND`, `OR`, `IRR`, `MIRR`, `XNPV`, `PROB`, `FORECAST`, `FORECAST.LINEAR`,
  `PERCENTILE.EXC`, `TRIMMEAN`), and `SEQUENCE`, which takes no range at all — its arguments are a size, a
  start and a step, so there is no rectangle to hand it and the sweep is blind to it for good.
  Those 22 have the roster and the by-hand list as their only defence, so the
  blind set is itself pinned by name and gaining a member fails the suite too.

**Not supported (by design).**

- A **dry cell** whose whole formula is an array is `#VALUE!` **unless the array came from one of the four
  [producers](#dynamic-array-producers)**, which answer their top-left element instead. The boundary rule is
  therefore per node kind, and there are three cases. (1) A **producer** — `=FILTER(...)`, `=SORT(...)`,
  `=UNIQUE(...)`, `=SEQUENCE(...)`, and any expression built over one — yields the array's top-left value,
  Excel's `@`-on-an-array rule, the same answer in every cell and in both of Excel's entry modes: see the
  table in that section. (2) A **range operand under an operator or a lifted function** keeps `#VALUE!` here —
  `=A1:A3*2`, `=LEN(A1:A3)`, `=ROUND(A1:A3,0)`, `=-A1:A3` — because the lift happens inside the *consumers*
  and the cell boundary is not one of them: the cell sees `LEN`'s ordinary scalar body handed a range. Excel
  does something different from either engine there, and it is a genuine gap rather than a rule: typed, Excel
  applies implicit intersection to each range operand *before* the operator, using the formula cell's own row,
  so with `A1:A3` = 5, 0, 9 a bare `=-A1:A3` is `-5` in `C1`, `0` in `C2`, `-9` in `C3` and `#VALUE!` in `C5`,
  and `=ROUND(A1:A3,0)` is 5, 0, 9 and `#VALUE!` in those same cells; array-entered it takes the top-left in
  every cell (`-5`, `5`). Both columns measured on Aspose.Cells 26.6.0, 2026-09-10. Today's `#VALUE!` is
  pinned so closing that gap has to be deliberate — and note that the pinning test's own comment still says
  the plain form is `#VALUE!` everywhere, which the measurement above contradicts for a formula row *inside*
  the range. (3) A bare `IF(range…)` or range comparison is `#VALUE!` for the same reason as (2) —
  `=IF(B2:B5="Show",1,0)` and `=IF(TRUE,A1:A3,B1)` on their own are errors — a known inconsistency with case
  (1) beside it. In every case, wrapping the expression in a consumer works: `=SUM(LEN(A1:A3))` in that same
  cell is `3` for `A1:A3` = 5, 0, 9 (one character each). Arrays still exist only as *arguments* and as a
  producer's collapsed top-left, never as a multi-cell cell value: the per-cell cache stays strictly scalar
  and there is no spill. This does **not** contradict
  [implicit intersection at the cell boundary](#implicit-intersection-at-the-cell-boundary): that rule
  intersects a *reference* — the bare `=A1:A3` beside these is `A3` in `C3` — while a computed array has no
  worksheet position, which is why the producer's answer is position-independent.
- **A broadcast product in a bare cell is `#VALUE!` too**, and the boundary is where that gap lives rather
  than in the broadcasting rule: `=A1:C3*E1:E3` typed into a cell never enters the element-wise evaluation,
  so it is the multiplication operator's own `#VALUE!`, while `=SUM(A1:C3*E1:E3)` in that same cell is 108.
  Excel's rule for the typed form is not an array rule at all — it applies implicit intersection to **each
  range operand separately, before the operator**, using the formula cell's own row and column, so the same
  formula answers differently in different cells: `=A1:A3*E1:E3` is **21** in `J3` (`A3`×`E3`) and **1** in
  `J1`, `=E1:E3*10` is **20** in `L2`, **30** in `L3` and `#VALUE!` in `L5` (row 5 misses `E1:E3`), and a 2-D
  operand never intersects, so `=A1:C3*E1:E3` is `#VALUE!` there as well — the only way to get the top-left
  element of the computed array is to ask for it, `=INDEX(E1:E3*10,1,1)` = **10** (all typed entry, measured
  on Aspose.Cells 26.6.0, 2026-09-10). Closing that belongs to the
  [cell boundary](#implicit-intersection-at-the-cell-boundary) rule — the per-operand behaviour above is what
  the array half of `@` has to be reconciled with — not to the consumers described here; today's `#VALUE!` is
  pinned by `CellBoundaryIntersectionTests` so the change has to be deliberate.
- **Unary `+` is deliberately not lifted.** It is Excel's reference-preserving no-op, so `+A1:A3` stays a
  *reference* and the consumer folds it on the ordinary range path: `SUM(+A1:A3)` = 6 for `A1:A3` = 1,2,3,
  exactly as `SUM(A1:A3)` does, and unchanged by the lift. The cost of keeping it opaque is that a `-` over
  it has nothing to lift: `SUM(-(+A1:A3))` is `#VALUE!` where Excel answers -6 (Aspose.Cells 26.6.0,
  `Ctrl+Shift+Enter`, measured 2026-09-09). Write `SUM(-A1:A3)` instead.
- A **reference-returning function** as `ROW`/`COLUMN`'s argument stays a scalar:
  `SUM(ROW(INDEX(A1:A3,1,1)))` is `1`, the top row of the resolved reference, not the vector `[1,2,3]`.
  Discovering its shape would resolve the argument a second time and draw a volatile twice, so the array
  shape is deliberately deferred there.
- The **criteria / positional-scan** family does not read a computed array *wherever it can recognise one* —
  it **rejects** it with `#REF!`. `SUMIF`/`SUMIFS`, `COUNTIF`/`COUNTIFS`, `AVERAGEIF`/`AVERAGEIFS` and
  `MAXIFS`/`MINIFS` walk their arguments position by position, and every range slot they take — criteria
  range and sum/average/max/min range alike — requires a *reference*: an argument that is not a reference
  node and that the element-wise evaluation would stream is refused before the scan opens, in every slot and
  at every arity. `COUNTIF(A1:A3*1,">0")`, `SUMIF(A1:A3*1,">0")`, `COUNTIFS(A1:A3*1,">0")`,
  `AVERAGEIF(A1:A3*1,">0")`, `SUMIFS(B1:B3,A1:A3*1,">0")`, `SUMIFS(A1:A3*1,B1:B3,">0")`,
  `SUMIF(A1:A3,">0",B1:B3*1)`, `COUNTIFS(A1:A3,">0",B1:B3*1,">1")`, `COUNTIF(ROW(A1:A3),">1")`,
  `COUNTIF(-A1:A3,"<0")` and `COUNTIF(LEN(A1:A3),">0")` are all `#REF!`. Excel refuses the family the same
  way: for a **computed** argument (`SUMIFS((A1:A3)*1,A1:A3,">0")` and its seven siblings) it answers
  `#VALUE!` on plain entry and `#REF!` array-entered, and for a dynamic-array **producer** in the slot it
  answers `#REF!` in *both* entry modes — `COUNTIF(FILTER(A1:A3,A1:A3>0),">5")`,
  `COUNTIF(SEQUENCE(5),">3")`, `SUMIF(SORT(A1:A3),">0")` and `COUNTIF(UNIQUE(A1:A3),">0")`, all measured on
  Aspose.Cells 26.6.0, 2026-09-10, and answered the same way here now that
  [the four exist](#dynamic-array-producers). `#REF!` is therefore both the array-entered answer this section
  reproduces and the one answer the two producer columns agree on, which is why the rule is `#REF!` and not
  `#VALUE!`. Pinned by `CriteriaComputedArgumentTests` and
  `MathAggregateTests.CriteriaFamily_RejectsAComputedArrayWithRef`. A **broadcast** or composite argument is
  the same rejection — the family never enters the element-wise evaluation — so `SUMIF(A1:C3*H1:H2,">0")`,
  `COUNTIF(A1:C3*H1:H2,">0")` and `SUMIFS(A1:C3,A1:C3*H1:H2,">0")` are `#REF!` here, matching the oracle's
  array-entered column (`#VALUE!` typed; measured 2026-09-10, pinned by
  `MiniCseConsumerTests.CriteriaFamily_OverABroadcastArray_IsRef`), and so is a **lifted** argument:
  `SUMIFS(LEN(A1:A3),A1:A3,">0")` is `#REF!`, with the same `#VALUE!`-plain / `#REF!`-array-entered split on
  the oracle (pinned by `MiniCseConsumerTests.CriteriaFamily_OverALiftedFunction_IsRef`). What is **not**
  rejected is whatever is already a reference or is not array-eligible: a reference-returning function
  (`CHOOSE`, `OFFSET`, `INDEX`), a defined name, a single cell and a whole column all stay ranges, so
  `COUNTIF(CHOOSE(1,A1:A3,B1:B3),">0")` and `COUNTIF(OFFSET(A1,0,0,3,1),">0")` are `2`, as on the oracle in
  both modes. Four shapes are **deliberate deviations**, each pinned as one in
  `CriteriaComputedArgumentTests` — three left for the compatibility sweep and the fourth the standing `LET`
  limit: `COUNTIF(IF(TRUE,A1:A3,B1:B3),">0")` is `0` here where the oracle answers `2` in *both*
  entry modes — a scalar-conditioned `IF` is an opaque scalar here rather than its
  branch's reference, and closing that is the sweep's own item, deliberately not part of this rule;
  `COUNTIF(5,">0")` and `COUNTIF(A1*1,">0")` are `1` where the oracle answers `#REF!` in both modes (a bare
  *scalar* in a range slot, a shape no array producer takes); and `SUMIF(A:A*1,">0")` is `0` where the oracle
  answers `#REF!` in both modes (the cost guard refuses a whole-column operand, so the argument is never
  array-eligible and the gate never sees it); and a `LET` is `0` from **either** side of its binding —
  `COUNTIF(LET(r,A1:A3,r*1),">0")` and `LET(r,A1:A3*1,COUNTIF(r,">0"))`, with the `SUMIF` twins alike — where
  the oracle answers `#REF!` array-entered (`#VALUE!` typed), because a `LET` node is an opaque scalar to the
  shape probe while a `LET`-bound name *is* a reference node whose binding was already collapsed when it was
  captured, so the gate's predicate sees no array either way. That last one is **pre-existing** (measured
  identical before the rule landed) and is a standing limit rather than part of this rule:
  `LET(f,FILTER(A1:A3,A1:A3>0),COUNTIF(f,">0"))` is exactly that shape and is the one deliberately failing
  pin in the suite — see the `LET` entry under
  [dynamic array producers](#dynamic-array-producers). `SUMPRODUCT` is the one member of that family that opted in to
  computed arrays — `SUMPRODUCT((A1:A3<>0)*1)` = 2 and `SUMPRODUCT(A1:A3*1,B1:B3)` = 32, matching the oracle
  in both modes — and the fold-based consumers listed under **Supported** above (`SUM(IF(…))` and friends)
  have always taken them. `SUBTOTAL` and AGGREGATE's reference form take neither path — they reject a
  computed array outright, a lifted one included (`SUBTOTAL(9,LEN(A1:A3))` and `AGGREGATE(9,4,LEN(A1:A3))`
  are `#VALUE!` on both engines); AGGREGATE's array form is the one that consumes it, lifts included
  (`AGGREGATE(15,6,LEN(A1:A3),1)` = 1, measured on both).
- An **open/whole-column** range in an array position is refused and the consumer stays on its ordinary
  scalar/range path — the one exception is the `INDEX(ROW($A:$A), n)` identity above, which returns `n`
  without materializing the column. `SMALL(IF(A:A=…, ROW(A:A)), k)` over an *open* column is therefore
  not array-evaluated. A lifted call over one is refused the same way, and the refusal is *tolerated* rather
  than fatal: the call collapses to a single opaque scalar evaluated once, so `SUM(LEN(A:A))` is `#VALUE!`
  (the scalar `LEN` of a range) while an enclosing array expression keeps working —
  `SUM(IF(A1:A3>0,1,LEN(B:B)))` is still 3. Excel folds the open column instead (`SUM(LEN(A:A))` = 3 over
  three one-character cells, Aspose.Cells 26.6.0 array-entered, 2026-09-09); a formula that works over
  `A1:A3` and is then dragged to a whole column gets the old `#VALUE!` back, with no other warning.
- A **scalar** condition keeps `IF`'s native short-circuit: only the taken branch is evaluated, and only an
  array *condition* drives the zip. The taken branch is still read as an array when it *is* one, though — a
  producer, a lifted call or an operator result — so `SUM(IF(TRUE,SEQUENCE(3),0))` is 6 and
  `ROWS(IF(TRUE,SEQUENCE(3),0))` is 3, matching Excel in *both* entry modes, and
  `SUM(IF(TRUE,A1:C3*2,0))` is 90 over `A1:C3` = 1…9, matching its array-entered column (`#VALUE!` typed).
  The exception is a branch that is a **bare range**: `SUM(IF(TRUE,A1:C3,0))` is `#VALUE!` here where Excel
  answers 45 in both modes, deliberately left alone because moving it would answer "does `IF` return a
  reference?" for that one shape while its siblings stay unanswered. It is recorded for the compatibility
  sweep and pinned as a gap. All measured on Aspose.Cells 26.6.0, 2026-09-10.

**Known divergences.** Each of these is pinned by a test as a *gap*, not asserted as Excel's rule, so closing
one is always a deliberate edit; the single entry with no pin says so in its own words. Excel here means
Aspose.Cells 26.6.0, the version this project measures against, with the formula array-entered
(`Ctrl+Shift+Enter`) — the entry form whose semantics this element-wise evaluation reproduces without the
keystroke — and any figure taken from the typed form is labelled *typed* where it appears.

- **Uncovered on BOTH axes at once is `#N/A` here, and the oracle has no answer to match.** The
  broadcasting rule above is per axis, so a 2x2 against a 3x3 leaves five of the nine positions uncovered on
  the row axis, the column axis or both, and each is `#N/A` while the four covered positions compute:
  `SUM(A1:B2*A1:C3)` is `#N/A` and `COUNT(A1:B2*A1:C3)` = 4. Excel answers `#N/A` too at the **one**-axis
  shortfalls measured here — `INDEX(A1:B3*A1:C2,3,1)`, a 3x2 against a 2x3, is `#N/A` on both engines
  (CSE-entered and typed alike), and so are the shapes pinned here in which the *vector* is the shorter
  operand (`SUM(A1:C3*H1:H2)` and `SUM(A1:C3*E5:F5)` are `#N/A` with `COUNT` 6 on both, CSE-entered) — but that
  agreement is **not** universal: a 2-D rectangle shorter than a ROW vector is a measured counterexample,
  the next entry below. For the doubly-uncovered shape it answers **`#REF!`** from `INDEX` — in typed *and*
  array-entered form alike — and its calculator never returns at all for `SUM` or `COUNT` over that same
  array (no answer after ten minutes here, and over 200 s in each entry mode when the phase first met it,
  while `ROWS`/`COLUMNS` over it still report 3 and 3 immediately). All
  measured on Aspose.Cells 26.6.0, 2026-09-10. That is why this one is *not* on the compatibility sweep:
  a `#REF!` on two axes against an `#N/A` on one is not a rule to copy, and a computation that does not
  terminate is not a behaviour to reproduce. Pinned by
  `VectorBroadcastingTests.UncoveredOnBothAxes_StaysNotAvailable_WhereTheOracleIsSelfInconsistent`, whose
  comment carries the measurement.
- **A 2-D RECTANGLE shorter than a ROW vector is `#N/A` here, where the oracle fills the uncovered column
  with `0`.** This is the one *one*-axis shortfall found so far on which the two engines disagree, and it is
  the direction Phase 10 never fixtured: the phase's other mismatched vector shapes all have the VECTOR as
  the shorter operand, so this class went untested. Over the broadcasting fixture (`A1:C3` = 1..9
  row-major, `E5:G5` = 10,20,30) a 3x2 rectangle against that 1x3 row takes a 3x3 extent whose third column
  the rectangle does not cover, and there `SUM(A1:B3*E5:G5)` is `#N/A` here against **420** on the oracle,
  `COUNT(A1:B3*E5:G5)` = 6 against **9**, and `INDEX(A1:B3*E5:G5,1,3)` is `#N/A` against **0**
  (Aspose.Cells 26.6.0, CSE-entered, measured 2026-09-10). The *extent* agrees — `INDEX(…,4,1)` and
  `INDEX(…,1,4)` are `#REF!` on both — and so does the covered arithmetic, since
  `SUM(IFERROR(A1:B3*E5:G5,0))` is 420 on both engines in that mode; the whole difference is what fills the
  uncovered column. MySheet keeps its `#N/A` because the oracle is not self-consistent there, all three
  checks CSE-entered: drop the rectangle to two rows and the aggregates go back to `#N/A` with
  `COUNT(A1:B2*E5:G5)` = 4, yet `INDEX(A1:B2*E5:G5,1,3)` is still **0** — a position `COUNT` therefore does
  not count; and the COLUMN-vector mirror never fills at all, since a 2x3 rectangle against the 3x1 `E1:E3`
  is `#N/A` with `COUNT` 6 and `INDEX(A1:C2*E1:E3,3,1)` `#N/A` on **both** engines, as is a 1x2 rectangle
  against `E5:G5` (`#N/A`, `COUNT` 2, on both). Recorded for the planned Excel-compatibility sweep. Pinned
  by `VectorBroadcastingTests.RectangleShorterThanARowVector_StaysNotAvailable_WhereTheOracleFillsWithZero`,
  whose comment carries every number above.
- **`SUM(ROW(Ghost!A1:A3))`** — a rectangle written *literally* on a sheet that does not exist, in an array
  position — answers `6`, the row numbers `1+2+3`, where Excel answers `#REF!`. The scalar
  `ROW(Ghost!A1:A3)` in the same workbook is already `#REF!`, and so is the array path over a name that
  stands for the same range (`SUM(ROW(GhostName))`): the divergence is only the written-out rectangle, whose
  syntactic fast path goes straight to a row/column vector and never resolves the reference, so the
  missing-sheet guard that every resolving path runs has nothing to run on. A *lifted* function over the same
  ghost rectangle does read cells and therefore does report `#REF!` (`SUM(LEN(Ghost!A1:A3))`), which is why
  the two adjacent shapes disagree. Pinned by
  `MiniCseConsumerTests.Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap`.
- **`INDEX(<computed array>, 0)`** is `#REF!` here, where Excel intersects the whole vector and answers its
  first element — `INDEX(LEN(A1:A3),0)` and `INDEX(ROW(A1:A3),0)` are both **1** there (measured 2026-09-09,
  plain and array-entered alike). MySheet rejects `row_num` or `column_num` below 1 outright.
- **An empty argument slot keeps the function's documented default**, where Excel reads it as a supplied `0`:
  `FIXED(A1,,TRUE)` is `1.00` here and **`1`** there, `DOLLAR(A1,)` is `$1.00` here and **`$1`** there, for
  `A1` = 1 (measured 2026-09-09). With the slot fully absent both engines agree — `FIXED(A1)` and
  `DOLLAR(A1)` are `1.00` and `$1.00` on each — so the divergence is the *empty* slot, not the default, and
  it applies equally to the scalar call and to the lifted `FIXED(A1:A3,,TRUE)`.
- **A lifted call under a unary `+` is not lifted either.** `+` is Excel's reference-preserving no-op and
  MySheet keeps the whole `+`-expression opaque, which hides what is *inside* it from the mini-CSE: over
  `A1:A3` = 1, 22, 333, `SUM(+LEN(A1:A3))` is `#VALUE!` here and **6** in Excel (measured 2026-09-09). It is
  the sibling of the `SUM(-(+A1:A3))` = -6 case above — the same opaque `+`, with a lifted *function* inside
  it instead of a unary operator — and `SUM(LEN(+A1:A3))`, the `+` on the inside, is the same `#VALUE!` here
  against the same **6** there. Write `SUM(LEN(A1:A3))`. Pinned by
  `ElementwiseLiftingTests.LiftedCall_UnderAnOpaqueUnaryPlus_IsNotLifted_KnownDivergence`.
- **A defined NAME in an array position is whatever it is bound to** — the rule itself is *agreement*, and
  what this entry records are the three shapes still refused. A name bound to a rectangle is array-eligible in
  a **nested** array position and answers exactly what the written-out rectangle answers. For a `MyName`
  bound to 1, 22, 333 and a `Rng` bound to `A1:A3` = 5, 0, 9, all measured on Aspose.Cells 26.6.0 array-entered
  (2026-09-10): `SUM(LEN(MyName))` = **6**, `SUM(-MyName)` = **-356**, `SUM(MyName%)` = **3.56**,
  `SUM(MyName*2)` = **712**, `SUM((MyName>1)*1)` = `SUM(IF(MyName>1,1,0))` = `SUMPRODUCT(--(MyName>1))` =
  **2**, `COUNT((Rng<>"")*1)` = `COUNT(Rng*1)` = **3**, `SUM((Rng<>0)*1)` = **2**,
  `SMALL(IF(Rng>0,Rng),1)` = **5** and `INDEX(Rng*2,3)` = **18** — the same values as the literal twins,
  which is the rule stated as a test. A name on a **missing sheet** streams the literal's per-element `#REF!`
  the same way: `SUM((GhostName<>0)*1)` is `#REF!` and `COUNT((GhostName<>"")*1)` is `0`, both modes on the
  oracle. Reading the name itself is unaffected (`SUM(MyName)` = 356, `SUM(ROW(MyName))` = 6 on both
  engines), and at a consumer's **top level** a bare name is still a *reference* that keeps the reference
  path, exactly as a bare literal range does — `SUBTOTAL(9,Rng)` = 14, `AGGREGATE(9,4,Rng)` = 14,
  `SUM(A1:INDEX(Rng,3))` = 14 and `ISREF(INDEX(Rng,2))` = `TRUE` — because that path carries what an
  element-wise stream cannot: the nested-`SUBTOTAL` skip, the engine's column-major first-error scan and a
  reference-returning `INDEX`. Pinned by `DefinedNameArrayEligibilityTests` and
  `ElementwiseLiftingTests.LiftedShapes_OverADefinedName_AreLifted`. Three shapes stay refused, each a
  deliberate deviation pinned as one: an **open-range** name meets the cost guard, so `SUM((MyCol<>0)*1)`
  with `MyCol` = `$A:$A` is `1` here — the truthy reference value — against the oracle's **2**
  array-entered (`0` typed); a **union** name resolves to a scalar, which makes the whole expression
  scalar-only, so `SUM((UnN<>0)*1)` is `1` against **2** array-entered (`#VALUE!` typed) and the *literal*
  union twin is `#VALUE!` here, making this the one row where a name does not match its literal; and a `LET`
  node in a consumer's own argument slot stays opaque because the shape probe does not look inside it, so
  `SUM(LET(r,Rng,(r<>0)*1))` is `1` against **2** in both entry modes — the same standing `LET` limit the
  [producers section](#dynamic-array-producers) records, still open. A `LET`-bound name *inside* an array
  position does resolve **when the name is bound to a range**,
  through the `LET` scope that [name resolution](#named-ranges) checks first:
  `LET(r,A1:A3,SUM((r<>0)*1))` = **2**, `LET(r,A1:A3,COUNT(r*1))` = **3** and
  `LET(r,A1:A3,INDEX(r*2,3))` = **18**, matching the oracle in both entry modes where they were `1`, `0` and
  `#REF!` before this rule. A name bound to a **computed array** does not: the binding is evaluated as a
  scalar when it is captured, so `LET(r,A1:A3*1,COUNT(r*1))` is `0`, `LET(r,A1:A3*1,SUM(r*1))` is `#VALUE!`
  and `LET(r,A1:A3*1,INDEX(r*2,3))` is `#REF!` against the oracle's **3**, **14** and **18** in both entry
  modes — unchanged by this rule (measured on Aspose.Cells 26.6.0, 2026-09-10, and on the engine before and
  after the rule), and it is the same standing `LET` limit — a producer bound by `LET` collapses for exactly
  this reason.
- **A range-aware function is never lifted over its SCALAR slots.** Excel lifts a range-aware function too:
  it consumes the range in the slot that takes one and repeats the *whole call* per element of a rectangle
  handed to any other slot. MySheet's classification is per *function*, not per slot, so a rectangle in a
  scalar slot stays a range and the call answers once. Over `A1:A3` = 1, 2, 3 and `B1:B3` = 10, 20, 30, with
  Excel's answer first and MySheet's in brackets, all measured 2026-09-09:
  `SUM(MATCH(A1:A3,A1:A3,0))` **6** [`#N/A`], `SUM(VLOOKUP(A1:A3,A1:B3,2,FALSE))` **60** [`#VALUE!`],
  `SUM(CHOOSE(A1:A3,10,20,30))` **60** [`#VALUE!`], `SUM(LARGE(A1:A3,A1:A3))` **6** [`#VALUE!`],
  `SUM(COUNTIF(A1:A3,A1:A3))` **3** [`0`], `SUM(INDEX(B1:B3,A1:A3))` **60** [`#VALUE!`],
  `SUM(RANK(A1:A3,A1:A3))` **6** [`#VALUE!`], `SUM(WORKDAY(A1:A3,1))` **9** [`#VALUE!`],
  `SUM(NETWORKDAYS.INTL(A1:A3,4))` **9** [`#VALUE!`], `SUM(NPV(A1:A3/10,10,20,30))` **120.92** [`#VALUE!`],
  `SUM(TYPE(A1:A3))` **3** [`16`], `SUM(RANDBETWEEN(A1:A3,A1:A3))` **6** [`#VALUE!`]. Two of MySheet's
  answers are **silent** rather than errors: `COUNTIF`'s `0` (the rectangle sits in its *criteria* slot, not
  in a range slot, so the `#REF!` rejection above does not reach it and the collapsed argument matches no
  criterion) and `TYPE`'s `16` (the type code of the `#VALUE!` it was handed). Plain
  `NETWORKDAYS` is the one member of the family Excel does *not* lift — `SUM(NETWORKDAYS(A1:A3,B1:B3))` is
  **8** there, which is `NETWORKDAYS(A1,B1)` alone, an implicit intersection to the first element rather than
  a per-element lift, and `#VALUE!` here. Pinned by
  `ElementwiseLiftingTests.AConsumesFunction_IsNotLiftedOverItsScalarSlots_KnownDivergence`.

Every one of these is deliberately left as it is for now, and all but two are recorded for a planned
Excel-compatibility sweep: the two-axis mismatch is not, because the oracle offers no answer to match there,
and neither is the ghost-sheet rectangle, whose cause is a syntactic fast path rather than a rule.

Volatile sub-expressions inside the array behave like any other volatile: a `RAND()` (broadcast, or in a
range cell the comparison reads) taints the consuming cell, so [`Recalculate()`](#the-epoch-model)
refreshes it while a non-volatile array formula stays cached.

## Named ranges

A workbook can define **names** that stand for an expression — usually a sheet-qualified range or cell,
but any expression (a constant, a formula, another name) is allowed. Names are workbook-level and
**case-insensitive**, exactly like Excel.

> A named range is **not** an Excel **Table** (a ListObject). A name is a static alias for one expression;
> a table is a named region with named columns, a totals row, a range that grows as rows are added, and its
> own reference syntax (`Tabela1[Valor]`, `[@Valor]`). MySheet models names *and* the table model
> ([Tables](#tables) — name, range, header/totals flags and column names), but neither the
> structured-reference syntax nor the self-growing range: a resize is a second `DefineTable` call. See
> [Excel interop → Scope and limitations](excel-interop.md#scope-and-limitations).

```csharp
var workbook = new Workbook();
var data = workbook.Sheets.Add("Data");
data["A1"] = new NumberValue(10);
data["A2"] = new NumberValue(20);
data["A3"] = new NumberValue(30);

// Convenience overload: parses the text. References MUST be sheet-qualified (names have no implicit
// sheet); a leading '=' and '$' markers are optional.
workbook.DefineName("Sales", "Data!A1:A3");

// Expression overload: a name can point at any expression, e.g. a constant.
workbook.DefineName("Rate", new NumberValue(0.1));

var main = workbook.Sheets.Add("Main");
ExpressionParser.Parse("=SUM(Sales)", main).Evaluate(workbook);   // 60
ExpressionParser.Parse("=Rate*100", main).Evaluate(workbook);     // 10
```

**Definition.** `Workbook.DefinedNames` is the `name → Expression` map. Define through
`DefineName(string, Expression)` or the `DefineName(string, string)` convenience overload, which parses
the text and **requires every reference to be sheet-qualified** — an unqualified reference (e.g. `A1:A3`)
throws `ArgumentException`, since a workbook-level name has no implicit sheet. An empty name, or one that
collides with a cell-reference shape (`A1`) or a boolean literal, is also rejected — and so is a name already
taken by a table, because names and tables share one namespace ([Tables](#tables)).

**Resolution order.** A `NameReference` resolves in this order:

1. **`LET` scope first** (shadowing) — a `LET` binding with the same name wins, so
   `LET(Sales, 5, Sales+1)` is `6`, not a sum over the range. A `LET` binding captures a range/union
   node as a **reference value** (like a defined name does), so `LET(r, A1:C9, SUM(r))` and
   `LET(hdr, Data!$1:$1, MATCH(x, hdr, 0))` see the cells; a single cell is bound by value.
2. **`Workbook.DefinedNames`** — the name's expression is evaluated. A range/union stays a *reference*
   value, so range-aware functions expand it (`SUM(Sales)`); a single cell or constant evaluates to its
   scalar. The functions that require a syntactic reference — `VLOOKUP`/`HLOOKUP` (table), `INDEX`,
   `OFFSET`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `AREAS`, `ISREF` — accept a name that stands for a range
   (e.g. `VLOOKUP(2, Sales, 2)`).
3. Otherwise `#NAME?`.

A name used **bare in a cell** (`=Sales`) is not an error either: the reference it stands for is
[implicitly intersected](#implicit-intersection-at-the-cell-boundary) with the formula cell's row and
column, so `=Sales` over `Data!A1:A3` shows `Data!A3` when it is typed in row 3.

**Cycles.** A name that refers to itself, directly or through a chain (`A → B → A`), is detected by a
thread-local guard and yields `#REF!` instead of overflowing the stack.

## Tables

A workbook can also register **tables** — the model behind Excel's `<table>` part (a ListObject): a named,
sheet-anchored rectangle with named columns. `Workbook.Tables` is the read-only `name → Table` registry and
`DefineTable` is its only writer.

> **What is modelled, and what is not.** The registry holds the table *model* — the name, the range, the
> header/totals flags and the column names — and it survives `Save`/`Load`. The structured-reference
> **syntax** is not implemented yet: `=SUM(Tabela1[Valor])` raises `ParseException: Unexpected character '['
> (at position 11).` (measured 2026-09-10 on the release that introduces the registry — the trailing period is
> part of the message), and `ExcelFile.Load` does not populate the registry from an xlsx `<table>` part either
> ([Excel interop → Scope and limitations](excel-interop.md#scope-and-limitations)). So nothing in the
> evaluator reads a table yet: you register one to keep the model through a round trip, and to give the
> reference syntax something to resolve against when it lands.

```csharp
var workbook = new Workbook();
workbook.Sheets.Add("Data");

// The A1 overload: `reference` is exactly the xlsx <table ref="…"> string, so it spans the WHOLE table.
workbook.DefineTable(
    "Tabela1", "Data", "C2:E9", ["Produto", "Qtd", "Total"],
    hasHeaderRow: true, hasTotalsRow: true);

var table = workbook.Tables["tabela1"];   // case-insensitive, like Excel

table.HeaderRow;     // 2  — FirstRow, because HasHeaderRow
table.FirstDataRow;  // 3
table.LastDataRow;   // 8
table.TotalsRow;     // 9  — LastRow, because HasTotalsRow
table.DataRowCount;  // 6
table.FirstColumn;   // 3  (C)
table.LastColumn;    // 5  (E) — derived from ColumnNames.Count

table.TryGetColumnIndex("QTD", out var ordinal);      // true, ordinal = 1
table.TryGetColumnRange("qtd", out var column, out var firstRow, out var lastRow);
                                                      // true, 4, 3, 8 — the [#Data] band only
```

**Geometry.** `FirstRow`/`LastRow` span the whole range **including** the header and totals rows, exactly as
the xlsx `ref` attribute does; `FirstColumn` is the leftmost sheet column and the last one follows from
`ColumnNames.Count`. Everything else — `HeaderRow`, `TotalsRow`, `FirstDataRow`, `LastDataRow`,
`DataRowCount`, `LastColumn` — is **derived**, so the header, data and totals bands can never contradict each
other, and none of it goes on the wire.

**Names.** Table names follow **Excel's table-name rule**, not the defined-name one: start with a letter or
`_`, then only letters, digits, `.` and `_`, at most 255 characters, and nothing Excel would read as a
reference — a cell inside the grid in A1 form (`A1`, `T1`, `XFD1048576`), the `R1C1` form (`R1C1`, `r2c3`),
or the reserved single letters `C`/`R` (either case). MySheet also rejects `TRUE`/`FALSE`, which its
tokenizer reads as booleans, and a backslash anywhere in the name (Excel permits a leading one, but the
tokenizer would never read it into an identifier). Excel's own default names **are** legal (`Tabela1`,
`Table1`: their letter run is past the three-letter column bound, so they are not cells), as is a
letters-then-digits name outside the grid (`XFE1`, `A1048577`).

**One namespace with defined names.** Excel's Name Manager keeps tables and names together and refuses a
duplicate; `Workbook` enforces that symmetrically — `DefineTable` throws `ArgumentException` for a name a
defined name already holds, and `DefineName` throws for one a table holds — so the invariant never depends on
which came first.

**Columns.** Column names are human-authored and are *not* subject to the name rule, but they must be
non-blank, at least one, and **unique ignoring case** (Excel resolves `Table1[col]` to `Table1[Col]`), which
is also how `TryGetColumnIndex`/`TryGetColumnRange` compare them. In the A1 overload the range's width must
equal `columnNames.Count`.

**Redefining replaces.** A second `DefineTable` under the same name replaces the entry — that is how a table
is resized or its columns changed — and the registry keeps exactly one entry per name.

**The sheet need not exist.** Registration never touches `Sheets`, exactly like a defined name: a table can
name a sheet that has not been added (or has been removed). A missing sheet is an evaluation-time concern —
a reference into one resolves to `#REF!` rather than throwing (see [`GetCellValue`](#workbook)).

**Zero data rows is legal.** A header-only table (`ref="A1:A1"` with a header row) is a valid model state,
not an error: `DataRowCount` is `0` and `TryGetColumnRange` returns `false` for a *known* column, leaving the
`#REF!`-or-empty decision to the caller instead of handing back an inverted range.

**Redefinition and stale values.** Neither `DefineTable` nor `DefineName` evicts anything from the
memoization cache — a (re)definition changes no cell — so a formula that already read the old definition
keeps serving its memoized value. Two ways to make the change observable:

- Call `InvalidateCache()` yourself: the whole cache goes and everything recomputes lazily.
- Or let the incremental recalculation engine do it. `Workbook.CreateRecalculationEngine()` returns a
  `RecalculationEngine` — a reverse dependency graph that, given the cells you report as edited, evicts only
  the affected cone instead of the whole cache. It tracks definition changes separately (they touch no sheet,
  so no sheet's structural version moves) and answers one with a **full** invalidation rather than a partial
  evict: the next `RecalculationEngine.Recalculate(...)` calls `InvalidateCache()` for you and reports
  `Mode = RecalculationMode.FullFallback` with `DirtyCellCount = -1`. Its plan-before-you-act counterpart,
  `RecalculationEngine.EstimateImpact(...)`, reports `RecommendFull = true` (`ConeSize = -1`) for the same
  reason and does **not** consume the signal, so the order of the two calls does not matter: estimating first
  still leaves the following `Recalculate` to evict.

> `RecalculationEngine.Recalculate(edited)` is **not** `Workbook.Recalculate()`. The workbook method
> refreshes only volatile cells and advances the [volatile epoch](#the-epoch-model); the engine method is the
> edit-driven incremental pass described above. The two names are unrelated beyond the word.

## Volatile functions

Five functions are **volatile** — their result is not fixed by the cells they read. Four depend on the
clock or a random draw: `NOW()`, `TODAY()`, `RAND()` and `RANDBETWEEN(bottom, top)`. The fifth,
`INDIRECT(ref_text, [a1])`, depends on *which cells it reads at all* — the reference is assembled from
text at evaluation time, so no static dependency is knowable. MySheet gives them Excel's two defining
behaviours — *recalculate on demand* and *contagious volatility* — without a dependency graph, through
an **epoch cache model**.

### The epoch model

Within one epoch a volatile is computed **once** and cached, so every `NOW()`/`TODAY()` in a pass agrees on
the same instant and a `RAND()` cell read twice returns the same value. A cell that touches a volatile —
directly (`=NOW()`) or transitively (`=A1+1` where `A1=NOW()`) — is cached **and marked**; the mark rides
the same thread-local propagation the cycle detector uses, so volatility spreads to dependents for free.

- **`Recalculate()`** advances the epoch: it drops **only** the marked (volatile-touched) cells and
  re-samples the clock/RNG, leaving every stable cell cached. Values refresh **lazily** — the next read
  recomputes them. This is the cheap "give me the current time / a new random draw" call.
- **`InvalidateCache()`** still clears **everything** (use it after editing cell inputs) and also resets the
  epoch.

```csharp
using Danfma.MySheet;
using Danfma.MySheet.Parsing;

var workbook = new Workbook();
var sheet = workbook.Sheets.Add("Sheet1");
sheet["A1"] = ExpressionParser.Parse("=NOW()", sheet);
sheet["B1"] = ExpressionParser.Parse("=A1+1", sheet);   // transitively volatile

ComputedValue first = workbook.GetCellValue("Sheet1", "A1");
ComputedValue again = workbook.GetCellValue("Sheet1", "A1");   // same epoch → identical

workbook.Recalculate();                                        // advance the epoch
ComputedValue later = workbook.GetCellValue("Sheet1", "A1");   // re-sampled → newer
ComputedValue b = workbook.GetCellValue("Sheet1", "B1");       // B1 refreshed too (contagion)
```

The clock is sampled **lazily** — on the first volatile read of an epoch, not inside `Recalculate()` — so
`NOW()` reflects the instant the value was actually produced, and nothing is sampled if no volatile is read.

### Injecting the clock and the RNG

`NOW`/`TODAY` read `Workbook.TimeProvider` (default `TimeProvider.System`) in **local time**, like Excel.
Assign any `TimeProvider` to freeze time for a batch, or to make tests deterministic regardless of the
machine's clock and zone. `RAND`/`RANDBETWEEN` draw from a persistent RNG; set `Workbook.RandomSeed` (an
`int?`) **before the first volatile read** to make the whole run reproducible, or leave it `null` (default)
for a clock-seeded RNG.

```csharp
workbook.TimeProvider = TimeProvider.System;   // the default; swap for a fake to control the clock
workbook.RandomSeed = 12345;                    // reproducible RAND/RANDBETWEEN
```

The RNG advances across epochs and is never re-seeded, so successive `Recalculate()` passes produce
different draws while a single cell stays stable within its epoch. Neither `TimeProvider` nor `RandomSeed`
is serialized (they are runtime configuration): a loaded workbook starts from `TimeProvider.System` and an
unseeded RNG.

### Limits (by design)

- **No per-cell refresh.** You can refresh *all* volatiles (`Recalculate()`), not a single one. Refreshing
  just `A1=NOW()` while leaving a cached `B1=A1+1` stale would need a reverse dependency graph, which the
  engine deliberately does not keep — so the coarse but correct refresh is the one offered.
- **`OFFSET` is not volatile.** Excel marks `OFFSET` volatile as a safety net for automatic recalculation;
  here invalidation is explicit, so marking it would needlessly taint half a sheet — a conscious divergence.

## From expression back to formula text

`FormulaWriter` is the inverse of the parser — it renders an expression as Excel formula text (without
the leading `=`), emitting the minimal parentheses that re-parse to the same tree:

```csharp
using Danfma.MySheet.Parsing;

var expression = ExpressionParser.Parse("=SUM(A1:A2)*Sheet2!B1", sheet);
string formula = expression.ToFormula(sheet.Name);   // "SUM(A1:A2)*Sheet2!B1"
```

The `contextSheetName` argument controls qualification: references on that sheet stay unqualified
(`A1`); references to other sheets are qualified (`Sheet2!A1`, quoted when the name requires it). This
is what the Excel exporter uses in `FormulaMode.Formulas` ([Excel interop](excel-interop.md)).

## See also

- [ComputedValue and errors](computed-value.md) — reading evaluation results.
- [Custom functions](custom-functions.md) — extending the function set.
- [Function reference](function-reference.md) — the 164 built-in functions.
