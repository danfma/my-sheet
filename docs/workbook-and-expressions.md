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
- **Syntax errors throw `ParseException`** (with a `Position` property pointing at the offending token).
  Built-in functions also validate their argument count at parse time — `=ROUND(1)` throws, just as
  Excel would reject it at entry.
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
| 7 | unary `-` `+` | Unary prefix binds tighter than `^`, so `-2^2` is `(-2)^2 = 4`, matching Excel. |
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
- A bare range used where a scalar is expected (e.g. `=A1:B2` alone) evaluates to `#VALUE!`, as in
  Excel; ranges are consumed by functions (`SUM`, `COUNT`, lookups, …).
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

**Supported.** The consumers are the numeric aggregators (`SUM`, `COUNT`, `AVERAGE`, `MIN`, `MAX`, and —
through the same fold — `SMALL`, `LARGE`, the percentiles) and `INDEX`. An argument is evaluated as an
array when it is a **closed-range** comparison (`B2:B5="Show"`), an `IF` whose condition is such an array
(with or without an else branch), or `ROW(range)`; scalars broadcast across the vector. A branch-less `IF`
yields a logical `FALSE` where the condition is false, and the aggregators ignore logicals/text (exactly
why `SMALL(IF(…))` skips the non-matching rows). The first per-element error wins, as in Excel.

**Not supported (by design).**

- A **dry cell** whose whole formula is the array keeps `#VALUE!` — `=IF(B2:B5="Show",1,0)` on its own is
  still an error. Arrays exist only as *arguments* inside the consumers above, never as a cell's value
  (the per-cell cache stays strictly scalar).
- An **open/whole-column** range in an array position is refused and the consumer stays on its ordinary
  scalar/range path — the one exception is the `INDEX(ROW($A:$A), n)` identity above, which returns `n`
  without materializing the column. `SMALL(IF(A:A=…, ROW(A:A)), k)` over an *open* column is therefore
  not array-evaluated.
- A **scalar** condition keeps `IF`'s native short-circuit — only an array condition drives the zip.

Volatile sub-expressions inside the array behave like any other volatile: a `RAND()` (broadcast, or in a
range cell the comparison reads) taints the consuming cell, so [`Recalculate()`](#the-epoch-model)
refreshes it while a non-volatile array formula stays cached.

## Named ranges

A workbook can define **names** that stand for an expression — usually a sheet-qualified range or cell,
but any expression (a constant, a formula, another name) is allowed. Names are workbook-level and
**case-insensitive**, exactly like Excel.

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
collides with a cell-reference shape (`A1`) or a boolean literal, is also rejected.

**Resolution order.** A `NameReference` resolves in this order:

1. **`LET` scope first** (shadowing) — a `LET` binding with the same name wins, so
   `LET(Sales, 5, Sales+1)` is `6`, not a sum over the range.
2. **`Workbook.DefinedNames`** — the name's expression is evaluated. A range/union stays a *reference*
   value, so range-aware functions expand it (`SUM(Sales)`); a single cell or constant evaluates to its
   scalar. The functions that require a syntactic reference — `VLOOKUP`/`HLOOKUP` (table), `INDEX`,
   `OFFSET`, `ROWS`, `COLUMNS`, `AREAS`, `ISREF` — accept a name that stands for a range (e.g.
   `VLOOKUP(2, Sales, 2)`).
3. Otherwise `#NAME?`.

**Cycles.** A name that refers to itself, directly or through a chain (`A → B → A`), is detected by a
thread-local guard and yields `#REF!` instead of overflowing the stack.

## Volatile functions

Four functions are **volatile** — their result depends on the clock or a random draw, not only on the cells
they read: `NOW()`, `TODAY()`, `RAND()` and `RANDBETWEEN(bottom, top)`. MySheet gives them Excel's two
defining behaviours — *recalculate on demand* and *contagious volatility* — without a dependency graph,
through an **epoch cache model**.

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
- **`INDIRECT` is not implemented** (resolving a reference from text is a separate feature).

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
