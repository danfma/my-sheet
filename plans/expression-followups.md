# Expression follow-ups: INDIRECT, OFFSET truncation, cross-sheet `:`

Three Excel-compatibility follow-ups left over from the dynamic-range-endpoints feature (PR #2):
implement `INDIRECT`, truncate `OFFSET` height/width like Excel, and make cross-sheet `:` ranges
yield `#REF!`. Excel is the acceptance bar; subtle expected values are derived from **Aspose.Cells**
(already referenced in `external/BenchExpressions`) and baked in as constants.

## Scope

- **In:** the three items below, in `Danfma.MySheet` + its test project.
- **Out (separate, later):** the memory backlog (per-workbook interning, CellStore streaming
  serialization, sparse arrays).

## For Future Agents
As work proceeds: mark checkboxes `- [x]`; when a phase is done set its status to `Complete`, write
its **Phase Summary**, run its **Verification Plan** and record the result. Fill **Final Recap** +
**Deployment Plan** when all phases complete.

**Test framework is TUnit** — `dotnet test` does NOT work on .NET 10. Core suite:
`dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/<Class>/*"`
(baseline before this work: **990** passing).

**Aspose oracle procedure:** to derive an Excel-verified constant, add a throwaway block to
`external/BenchExpressions/Program.cs` that builds the same grid in `Aspose.Cells`, sets the formula,
calls `CalculateFormula()`, prints the result; copy the value into the test as a literal; remove the
throwaway (that tree is gitignored). No runtime Aspose dependency in `Danfma.MySheet.Tests`.

**Confirmed facts:**
- `INDIRECT` is not registered (`INDIRECT(...)` currently parses to a generic `FunctionCall` → `#NAME?`).
- Volatility mechanism: override `Expression.IsVolatile => true` and call
  `context.Workbook.MarkVolatileTouched()` in `Evaluate` (pattern: `Expressions/Dates/VolatileClock.cs`).
- Next free `MemoryPackUnion` tag on `Expression` is **318** (max used = 317 = `DynamicRange`).
- Runtime parse: `ExpressionParser.Parse(string expression, Sheet sheet)`.
- `EvaluationContext` carries `SheetName` (the current sheet; may be null for a context-less eval).

---

## Phase 1: INDIRECT function
Status: Complete

`INDIRECT(ref_text, [a1])` returns the reference named by the text. MySheet supports A1 style only.

- [ ] Create `Danfma.MySheet/Expressions/Lookup/Indirect.cs`:

  ```csharp
  using Danfma.MySheet.Parsing;
  using MemoryPack;

  namespace Danfma.MySheet.Expressions.Lookup;

  /// <summary>
  /// INDIRECT(ref_text, [a1]) — resolves the A1-style reference named by ref_text at evaluation time.
  /// Volatile (Excel parity). Only A1 style is supported; a1 = FALSE (R1C1) yields #REF!. Invalid text
  /// or an unknown sheet yields #REF!. A single-cell result is dereferenced to its value; a multi-cell
  /// result stays a reference (the OFFSET/CHOOSE convention), so range consumers expand it and it works
  /// as a ':' endpoint.
  /// </summary>
  [MemoryPackable]
  public sealed partial record Indirect(Expression[] Arguments) : Function
  {
      public override bool IsVolatile => true;

      public override ComputedValue Evaluate(EvaluationContext context)
      {
          context.Workbook.MarkVolatileTouched();

          if (!TryResolveReference(context, out var reference))
          {
              return ComputedValue.Error(Error.Ref);
          }

          return reference is CellReference cell
              ? context.Workbook.GetCellValue(cell.SheetName, cell.Id)   // single cell: dereference
              : ComputedValue.Reference(reference!);                      // range: stays a reference
      }

      public override bool TryResolveReference(EvaluationContext context, out Reference? reference)
      {
          reference = null;

          // a1 flag: default TRUE (A1). FALSE (R1C1) is unsupported → #REF!.
          if (Arguments.Length >= 2)
          {
              if (Arguments[1].Evaluate(context).CoerceToBoolean(out var a1) is not null || !a1)
              {
                  return false;
              }
          }

          if (!Arguments[0].Evaluate(context).TryGetText(out var refText))
          {
              return false; // ref_text must be text
          }

          if (context.SheetName is not { } sheetName
              || !context.Workbook.Sheets.TryGetValue(sheetName, out var sheet))
          {
              return false; // no current sheet to resolve unqualified refs against
          }

          Expression parsed;
          try
          {
              parsed = ExpressionParser.Parse("=" + refText, sheet);
          }
          catch (ParseException)
          {
              return false; // malformed reference text → #REF!
          }

          return parsed.TryResolveReference(context, out reference);
      }
  }
  ```

  > Confirm the exact `ComputedValue` accessors: the text extractor (`TryGetText`) and boolean coercion
  > (`CoerceToBoolean` returning `Error?`). Match whatever `Choose`/`Offset` use for arg coercion; adjust
  > the calls above to the real API if the names differ.

- [ ] Register in `Danfma.MySheet/Parsing/Parser.cs` (alongside the other lookup registrations, e.g. near `"OFFSET"`): `["INDIRECT"] = new(1, 2, arguments => new Indirect(arguments)),`
- [ ] Add the union tag in `Danfma.MySheet/Expressions/Expression.cs` after `[MemoryPackUnion(317, typeof(DynamicRange))]`: `[MemoryPackUnion(318, typeof(Indirect))]`
- [ ] Create `tests/Danfma.MySheet.Tests/Expressions/IndirectTests.cs` with the tests below.

### Verification Plan
- New tests (Aspose-verify the numeric expectations): `INDIRECT("A1")` with A1=5 → 5; `INDIRECT("Sheet2!B2")` with Sheet2!B2=7 → 7; `SUM(INDIRECT("A1:A3"))` over 1,2,3 → 6; `INDIRECT("A1"):A3` as a `:` endpoint spans A1:A3; `INDIRECT("not a ref")` → `#REF!`; `INDIRECT("A1", FALSE)` → `#REF!`; MemoryPack round-trip of an `Indirect` node.
- `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/IndirectTests/*"` → all pass.
- Full core suite: `... -c Release` → 990 baseline + new tests, 0 failures.

### Phase Summary
Done. Added `Danfma.MySheet/Expressions/Lookup/Indirect.cs` — `Indirect(Expression[]) : Function`,
volatile (`IsVolatile => true` + `MarkVolatileTouched()`), resolving `ref_text` at eval time by
parsing `"=" + refText` on the current sheet (`ExpressionParser.Parse`) and delegating to the parsed
node's `TryResolveReference`. Single-cell → dereferenced; range → `ComputedValue.Reference` (works as a
`:` endpoint). `a1 = FALSE` (R1C1), non-text arg, malformed text, or unknown sheet → `#REF!`. The `a1`
flag uses `CoerceToNumber` (== 0 → R1C1) since there is no `CoerceToBoolean`. Registered `"INDIRECT"`
(1-2 args) in `Parser.cs`; added `[MemoryPackUnion(318, typeof(Lookup.Indirect))]` in `Expression.cs`.
**Verification:** `IndirectTests` 7/7 (A1, qualified sheet, range-sum, `:`-endpoint, invalid→#REF!,
R1C1→#REF!, round-trip); full core suite 997/997 (990 + 7). Values self-evident (no Aspose needed).

---

## Phase 2: OFFSET truncates height/width (Excel parity)
Status: Complete

`Offset.cs` already truncates `rows`/`columns` via `(int)`, but the single-cell decision
`height == 1 && width == 1` (`Offset.cs:26` and `:52`) compares the raw doubles, so
`OFFSET(A1,0,0,1.5)` takes the multi-cell branch. Excel truncates height/width to integers.

- [ ] **Aspose oracle first:** confirm Excel's behavior for `=OFFSET(A1,0,0,1.9,1)` (and `=SUM(OFFSET(A1,0,0,1.9,1))` over a small grid) — expected: truncates 1.9→1, i.e. a single cell / 1-row range, identical to `OFFSET(A1,0,0,1,1)`. Record the exact result.
- [ ] In `Offset.cs`, truncate height/width once (toward zero, matching `(int)rows`/`(int)columns`) right after coercion, and base BOTH the single-cell check and `BuildRange` on the truncated ints. Concretely, in `TryComputeTarget` change the out-params for height/width to `int` computed as `(int)h`/`(int)w`, and have `Evaluate`/`TryResolveReference` test `height == 1 && width == 1` on those ints. (This is the intentional, Excel-correct version of the truncation that was previously reverted to preserve legacy double-compare behavior.)
- [ ] Build: `dotnet build Danfma.MySheet/Danfma.MySheet.csproj -c Release` — 0 warnings.

### Verification Plan
- New test `Offset_FractionalHeightWidth_TruncatesLikeExcel`: on a grid, `=SUM(OFFSET(A1,0,0,1.9,1))` equals the Aspose-verified value (a single cell / 1-row sum), NOT a 2-row sum.
- Existing OFFSET tests unchanged and green (they use integer height/width, so identical): `dotnet run ... -- --treenode-filter "/*/*/TryResolveReferenceTests/*"` and the OFFSET suite.
- Full core suite green.

### Phase Summary
Done. Aspose oracle confirmed Excel truncates OFFSET height/width toward zero: `=SUM(OFFSET(A1,0,0,1.9,1))`
over A1=10,A2=20 returns **10** (single cell), not 30. Changed `Offset.TryComputeTarget` to output
`int height, int width` (coerced to double then `(int)`-truncated, matching `(int)rows`/`(int)columns`);
`Evaluate`/`TryResolveReference` now test `height == 1 && width == 1` on the truncated ints and `BuildRange`
takes ints. **Verification:** new `Offset_FractionalHeightWidth_TruncatesLikeExcel` (SUM==10, Aspose-verified);
existing OFFSET tests unchanged (integer args → identical); full core suite 998/998.

---

## Phase 3: Cross-sheet `:` yields #REF!
Status: Not started

Two sub-cases (see the parser analysis): the **dynamic** case is the real silent bug; the **static
both-qualified** case already errors (a `ParseException`) and only differs from Excel cosmetically.

### 3a — Dynamic (safe, primary)
`DynamicRange.TryResolveReference` (`DynamicRange.cs`) takes the sheet from the Start endpoint and
ignores the End endpoint's sheet, so `INDEX(Sheet1!…):Sheet2!B2` silently resolves against Sheet1.

- [ ] In `DynamicRange.TryResolveReference`, after resolving `start` and `end`, extract BOTH endpoint sheets (a `CellReference`/`RangeReference` carries `SheetName`) and, if they differ (`!string.Equals(startSheet, endSheet, StringComparison.OrdinalIgnoreCase)`), set `reference = null; return false;` (→ `#REF!`). Same-sheet (the common `INDEX(...):$D1` case) is unaffected.
- [ ] Rewrite the existing `CountFamily_DynamicRangeOverMissingSheet_IsRefError` test (in `DynamicRangeEvaluationTests.cs`) so BOTH endpoints are explicitly on the missing sheet — `=COUNT(INDEX(Sheet2!A1:A3,2):Sheet2!A5)` — so it exercises the MISSING-sheet guard (same sheet, absent), not the new cross-sheet guard. Confirm it still yields `#REF!`.

### 3b — Static both-qualified (riskier, attempt then decide)
`Sheet1!A1:Sheet2!B2` is parsed inside `ParseQualifiedReference` (`Parser.cs:589`) and currently
throws `ParseException` because the second endpoint (`Sheet2`) is not a bare cell token. Excel yields
`#REF!` (a value error), which is friendlier for loading a workbook that contains such a formula.

- [ ] In `ParseQualifiedReference`'s colon branch, detect when the second endpoint is itself a qualified reference (a `!` follows it) targeting a DIFFERENT sheet, and build a `DynamicRange(firstRef, secondQualifiedRef)` instead of throwing — so evaluation runs the 3a cross-sheet guard and yields `#REF!`. The same-sheet qualified range (`Sheet1!A1:B2`, `Sheet1!A1:Sheet1!B2`) MUST keep its current `RangeReference`/`OpenRangeReference` fast path unchanged.
- [ ] **If this destabilizes the qualified-range parsing** (any existing parser/interop test regresses that cannot be reconciled), STOP, revert 3b only, and record in the Phase Summary that the static both-qualified case remains a `ParseException` (a tracked, acceptable divergence). 3a is the required part.

### Verification Plan
- New test `DynamicRange_CrossSheetEndpoints_IsRefError`: `=INDEX(Sheet2!A1:A3,2):A5` (Start on Sheet2, End on the current sheet) → `#REF!`; and a same-sheet control (`INDEX(A1:A3,2):A5`) still resolves.
- 3b (if landed): `Sheet1!A1:Sheet2!B2` used in a formula evaluates to `#REF!` (not a thrown `ParseException`); `Sheet1!A1:B2` still parses to a normal `RangeReference` on Sheet1.
- The rewritten missing-sheet test still `#REF!`.
- Full core suite green (990 + new).

### Phase Summary
_(write when phase completes)_

---

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
