# Design: Dynamic range endpoints for the `:` operator

- **Date:** 2026-07-07
- **Status:** Approved (pending spec review)
- **Area:** `Danfma.MySheet` — parser + expression evaluation
- **Related:** `plans/expression-parser.md`, `plans/fix-missing-sheet-ref.md`

## Problem

MySheet's parser rejects any formula where an endpoint of the `:` range operator
is a reference-returning **function** instead of a plain cell token. Real-world
Excel workbooks use this idiom for dynamic (expanding) ranges.

Confirmed repros (via `external/AllocProbe` `parse` mode):

```
THROW =INDEX($D:$D,2):$D1                      ParseException: The ':' range operator requires cell references (pos 14)
THROW =INDEX($D:$D,MATCH($C2,$D:$D,0)-1):$D1   ParseException (pos 33)
THROW =OFFSET(A1,1,0):B2                        ParseException (pos 14)
THROW =SUM(INDEX($D:$D,2):$D1)                  ParseException (pos 18)
OK    =A1:B2                                    RangeReference        (static ranges unaffected)
OK    =$D:$D                                    OpenRangeReference
```

Root cause: `Parser.ParseRange` (`Parsing/Parser.cs:448-466`) only builds a range
when both endpoints are static — a `CellReference` pair, or `TryEndpoint`-recognised
column/row tokens. A function node (`Index`, `Offset`, …) hits the `default` branch of
`TryEndpoint` and the method throws.

The formula that surfaced this comes from a production workbook loaded by the
`MySheetProcessor` pipeline (sheet `sheet1`), **not** the K1 benchmark corpus
(0 / 549,266 K1 formulas hit the pattern). Impact is real-world Excel input, not the
benchmark.

### Why "evaluate the endpoint and unwrap its reference" does NOT work

MySheet's reference-returning functions **dereference a single cell to its value**.
`Offset.cs:51-56`: a `1×1` OFFSET returns `GetCellValue(...)` (a value), and only a
multi-cell result returns `ComputedValue.Reference`. So evaluating `INDEX($D:$D,2)`
yields the *value* in `$D2`, losing the address. The design therefore needs a
resolution path that produces the target **reference without dereferencing**.

## Goals

- `ref1:ref2` parses and evaluates correctly when either endpoint is a
  reference-returning expression: `INDEX`, `OFFSET`, `CHOOSE`, `INDIRECT`, or a
  parenthesised reference.
- Downstream consumers that take a range (`COUNTIF`, `SUM`, `INDEX`, …) see the
  resolved concrete range — the real formula is
  `COUNTIF(INDEX($D:$D,MATCH(...)-1):$D1, $C2)`.
- Zero regression to static ranges (`A1:B2`, `$D:$D`, `1:5`) and to the K1 hot idiom
  `INDEX(ROW($A:$A),k)` (value path, 131k cells).

## Non-goals

- No change to how any function behaves in a value context.
- No destructive `.mysheet` format version bump (the new node is additive).
- Full R1C1 or structured-reference support.

## Design

### Component 1 — `TryResolveReference`: resolve an expression to a reference *without dereferencing*

A capability that returns an expression's **target reference** (address), distinct
from `Evaluate` (which may dereference a single cell to its value).

**Decision (approved):** a virtual method on `Expression`, each node owning its own
logic, mirroring the existing `Evaluate` pattern.

```csharp
public virtual bool TryResolveReference(EvaluationContext context, out Reference? reference)
{
    reference = null;   // default: not a reference-producing expression
    return false;
}
```

Overrides:

| Node | Resolves to |
| --- | --- |
| `CellReference`, `RangeReference`, `OpenRangeReference`, `UnionReference` | itself |
| `NameReference` | existing name resolution (delegates to `NamedReferences`) |
| `Index` | **concrete-range path only**: the target cell/range. Computed-array paths (`IndexIntoArray`, `IndexIntoOpenRowNumbers`) return **false** (no address — `#REF!`, as in Excel) |
| `Offset` | its target cell/range (refactor: compute the target reference first, then `Evaluate` dereferences a `1×1`) |
| `Choose` | resolves the chosen argument to a reference |
| `Indirect` | builds a reference from its string argument |
| parenthesised expression | transparent (delegates to the inner expression) |
| `DynamicRange` (Component 2) | the spanned range it resolves to |

This **subsumes and extends** today's `NamedReferences.TryResolveRaw` (which handles
only static `Reference` + `NameReference`). Existing callers
(`NamedReferences.TryResolveReference`, used by VLOOKUP/INDEX-base/OFFSET-base/ROWS/
COLUMNS) route through the same capability, so a function endpoint now resolves where
it previously failed — with the workbook's evaluation cycle guard reused to prevent
runaway recursion.

### Component 2 — `DynamicRange` AST node

```csharp
[MemoryPackable]
public sealed partial record DynamicRange(Expression Start, Expression End) : Reference;
```

- `TryResolveReference`: resolve `Start` and `End` to references (Component 1); take
  the bounding box of each; span from the min corner to the max corner; build a
  `RangeReference` (or `OpenRangeReference` if a bound is open) on `Start`'s sheet
  (mirrors `Parser.SheetOf`). Either endpoint failing to resolve → `#REF!`.
- `Evaluate`: resolve to the concrete range, then delegate to that range's `Evaluate`
  (so it inherits all range/scalar-context behaviour — implicit intersection, `#VALUE!`,
  etc. — for free).

Because `DynamicRange` itself implements `TryResolveReference`, range-consuming
functions (`COUNTIF`, `SUM`, `INDEX`) resolve it transparently to the concrete range.

### Component 3 — parser change

`Parser.ParseRange` (`Parsing/Parser.cs:448-466`): keep the static fast paths (the
`CellReference` pair and `TryBuildOpenRange`) exactly as they are. Replace the final
`throw` with:

```csharp
return new DynamicRange(left, right);
```

Static ranges are untouched — zero regression. Only the previously-throwing case now
produces a node.

### Serialization

`DynamicRange` is a new `[MemoryPackable]` reference node, additive to the reference
union. Existing `.mysheet` files are unaffected (they contain no such node); new files
that contain a dynamic range are forward-only (older MySheet cannot read them). No
destructive version bump.

## Risk assessment

- **Value paths unchanged.** Each function's `Evaluate` keeps identical output
  (`1×1` → value, multi-cell → reference). `TryResolveReference` is a *new, separate*
  method; refactors that split "compute target" from "dereference" must preserve
  `Evaluate`'s existing result.
- **K1 hot idiom safe.** `INDEX(ROW($A:$A),k)` uses the computed-array path →
  `TryResolveReference` returns false, `Evaluate` unchanged. The 131k-cell perf idiom
  is not touched.
- **Recursion.** Resolving a function endpoint may evaluate sub-expressions; reuse the
  existing evaluation cycle guard so a self-referential range degrades to `#REF!`.

## Excel compatibility (acceptance criterion)

Excel behaviour is the acceptance bar. Where the correct result is subtle, it is not
taken from memory: **`Aspose.Cells` (already referenced in `external/BenchExpressions`)
is an Excel-compatible engine and is used as the oracle** — the same formula is computed
in Aspose to derive the expected value, which is then baked as a constant into the TDD
test (so the committed test has no runtime Aspose dependency). Corners pinned this way:

1. **Corner rule** — `ref1:ref2` = smallest rectangle bounding both references; verify
   `B2:A1 == A1:B2` and mixed cell/range endpoints.
2. **Array-form INDEX as an endpoint** — e.g. `INDEX({1,2,3},2):A1`; confirm Excel's
   `#REF!` (our array path returns false → `#REF!`).
3. **Cross-sheet `:`** — `Sheet1!A1:Sheet2!B2`; match Excel's error behaviour rather than
   silently using `Start`'s sheet.

## Testing (TDD)

RED first, from the confirmed repros, then semantics, then regression.

- **Parse:** `INDEX(...):$D1`, `OFFSET(...):B2`, `$C2:INDEX(...)`, `CHOOSE(...):x`,
  `INDIRECT("D1"):x`, `(expr):x` all parse to `DynamicRange` instead of throwing.
- **Evaluate semantics:** the spanned range covers the correct cells and reads them —
  e.g. a small grid where `INDEX($D:$D,2):$D4` sums `D2:D4`; `COUNTIF(INDEX(...):$D1, k)`
  counts over the resolved range.
- **Reference-without-deref:** `INDEX($D:$D,2)` resolves to a reference to `$D2` (not
  its value) when used as an endpoint.
- **Regression:** `A1:B2`, `$D:$D`, `1:5` unchanged; `INDEX(ROW($A:$A),k)` still returns
  the row number (value); the K1 allocation/perf idiom unaffected (spot-check via
  `external/AllocProbe`).
- **Errors:** an endpoint that resolves to no reference (array INDEX, unbound name) →
  `#REF!`; cross-sheet endpoints use `Start`'s sheet.

## Decisions

- **Function coverage (approved):** all four — `INDEX`, `OFFSET`, `CHOOSE`, `INDIRECT`
  — plus parenthesised references, in one change.
- **Resolution style (approved):** virtual `TryResolveReference` on `Expression`, not a
  centralised type-switch.

## Open questions (confirm during implementation)

1. **Corner rule.** Excel's `ref1:ref2` is the smallest rectangle bounding both
   references (min top-left, max bottom-right). Adopt that; revisit only if a fixture
   disagrees.
