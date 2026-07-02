# Migrating to 2.0

MySheet 2.0 is a **namespace-only** reorganization of the AST. No formula behavior changed, no API was
renamed or removed, and **serialized workbooks are 100% compatible** in both directions of the type
move — the break is strictly compile-time: your `using` directives.

## Why

`Danfma.MySheet.Expressions` had grown to ~190 public types in one flat namespace. 2.0 splits it along
the same category lines as the [function reference](function-reference.md) (which mirrors Excel's own
catalog), and promotes the two evaluation *result* types to the root namespace, next to `Workbook`.

## Type → namespace map

| Types | 1.x namespace | 2.0 namespace |
| --- | --- | --- |
| `ComputedValue`, `ComputedValueKind`, `Error` | `Danfma.MySheet.Expressions` | `Danfma.MySheet` |
| Core AST: `Expression`, `ValueExpression`, `Function`, `Reference`, `FunctionCall`, `EvaluationContext`, `NumberValue`, `StringValue`, `BooleanValue`, `BlankValue`, `ErrorValue`, `CellReference`, `RangeReference`, `UnionReference`, `NameReference`, `BinaryOperation`, `UnaryOperation` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions` (unchanged) |
| Logical functions (12): `If`, `Ifs`, `Switch`, `And`, `Or`, `Xor`, `Not`, `IfError`, `IfNa`, `Let`, `TrueFunction`, `FalseFunction` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Logical` |
| Math & trigonometry functions (67): `Sum`, `SumIf`, `SumIfs`, `Product`, `Abs`, `Int`, `Round`/`RoundUp`/`RoundDown`, `Sqrt`, `Power`, `Mod`, `Sin` … `Acoth`, `Ceiling*`/`Floor*`, `Fact`, `Combin`, `Base`, `Roman`, … | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Mathematics` |
| Statistical functions (8): `Average`, `Count`, `CountA`, `CountBlank`, `CountIf`, `CountIfs`, `Max`, `Min` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Statistical` |
| Text functions (34): `Upper`, `Lower`, `Trim`, `Len`, `Left`, `Mid`, `Right`, `Value`, `Text`, `T`, `Concat`, `Concatenate`, `TextJoin`, `Find`, `Search`, `Substitute`, `RegexTest`/`RegexExtract`/`RegexReplace`, `Fixed`, `Dollar`, … | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Text` |
| Information functions (18): `IsNumber`, `IsBlank`, `IsError` and the `Is*` family, `N`, `Na`, `TypeFunction`, `ErrorType`, `SheetNumber`, `SheetsCount` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Information` |
| Lookup & reference functions (16): `VLookup`, `HLookup`, `XLookup`, `Lookup`, `Match`, `XMatch`, `Index`, `Choose`, `Offset`, `Row`, `Rows`, `Column`, `Columns`, `Address`, `Areas`, `FormulaText` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Lookup` |
| Financial functions (9): `Pmt`, `Pv`, `Fv`, `Nper`, `Ipmt`, `Ppmt`, `Npv`, `Rate`, `Irr` | `Danfma.MySheet.Expressions` | `Danfma.MySheet.Expressions.Financial` |

Folders mirror namespaces, so a type's source file is always where its namespace says.

Notes on the shape:

- A function's category is its section in the [function reference](function-reference.md), which mirrors
  Excel's catalog. That is why `SUM`/`SUMIF(S)` are **Mathematics** while `AVERAGE`/`COUNT*`/`MAX`/`MIN`
  are **Statistical**, and why `T` is **Text** (Excel lists it there) even though it shipped alongside
  the information functions.
- The namespace is `Mathematics`, not `Math`, so `Math.Sqrt(…)` keeps resolving to `System.Math` inside
  and around those types.
- `FunctionCall` (the custom-function AST node) stays in the core `Danfma.MySheet.Expressions`: it is
  the extensibility node, not a category.

## Updating your code

Most call sites only parse and evaluate — they touch `Workbook`, `ExpressionParser`, `ComputedValue`
and the value nodes. For those, nothing changes if you already import `Danfma.MySheet` (where
`Workbook` lives — `ComputedValue` and `Error` now come with it):

```csharp
// 1.x
using Danfma.MySheet;             // Workbook, Sheet
using Danfma.MySheet.Expressions; // NumberValue, ComputedValue, Error
using Danfma.MySheet.Parsing;     // ExpressionParser

// 2.0 — same directives still compile; ComputedValue/Error simply come from Danfma.MySheet now.
using Danfma.MySheet;             // Workbook, Sheet, ComputedValue, ComputedValueKind, Error
using Danfma.MySheet.Expressions; // Expression, NumberValue, FunctionCall, …
using Danfma.MySheet.Parsing;     // ExpressionParser, FormulaWriter
```

Code that names function records directly (building trees by hand, pattern-matching on nodes) adds the
category namespaces it uses:

```csharp
// 1.x
using Danfma.MySheet.Expressions;

Expression tree = new Sum([new CellReference("A1", "Sheet1"), new NumberValue(2)]);
bool isLookup = tree is VLookup or XLookup;

// 2.0
using Danfma.MySheet.Expressions;             // CellReference, NumberValue
using Danfma.MySheet.Expressions.Lookup;      // VLookup, XLookup
using Danfma.MySheet.Expressions.Mathematics; // Sum

Expression tree = new Sum([new CellReference("A1", "Sheet1"), new NumberValue(2)]);
bool isLookup = tree is VLookup or XLookup;
```

Watch out for two name collisions once you import the new namespaces:

- `Danfma.MySheet.Expressions.Lookup.Index` vs `System.Index` — qualify (`Lookup.Index`) or alias
  (`using IndexFunction = Danfma.MySheet.Expressions.Lookup.Index;`) where both are in scope.
- `Text` and `Lookup` are both a namespace **and** a record inside it (`…Text.Text` is the `TEXT`
  function, `…Lookup.Lookup` is `LOOKUP`); code sitting inside the `Danfma.MySheet.Expressions`
  namespace itself must qualify those two records.

## Serialization compatibility

The MemoryPack format is untouched. Expression nodes serialize through a tag-based
`[MemoryPackUnion]`, and **no tag changed** — namespaces are not part of the wire format. Workbooks
saved by 1.x load in 2.0 (and vice versa) byte-for-byte. This is guarded by a regression test: a binary
fixture serialized by the 1.x layout (`tests/Danfma.MySheet.Tests/Fixtures/workbook-pre-namespaces.msgpack.bin`)
is loaded and re-evaluated on every run.

## Behavior

None of the 164 built-in functions changed semantics, the parser and `FormulaWriter` accept and emit
exactly the same text, and evaluation results are identical. If a 1.x → 2.0 upgrade changes any computed
value, that is a bug — please report it.
