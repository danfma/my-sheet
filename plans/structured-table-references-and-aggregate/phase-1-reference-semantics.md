# Phase 1: ROW/COLUMN over any reference, Excel's implicit intersection, array-native SUMPRODUCT

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `reference-semantics`. Design dependencies: none. Adversarial verifier verdict: **needs-revision** (2 blockers, 2 majors, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Design decision

The three fixes land as one phase because they share one primitive: `NamedReferences.TryResolveReference` /
`CaptureValue` already turn "any node that denotes a reference" into a concrete `Reference`, and every fix is
a call to it at a place that currently type-switches syntactically. FIX A adds a shared `ReferencePosition`
helper (a new file, because `OrderSelection`-style `file static class` is invisible across Row.cs and
LookupFunctions.cs) as the terminal fallback of ROW/COLUMN, with `boundOpenRanges:false` so `ROW(A:A)` is the
DECLARED top row 1, not the populated one — the same choice Rows.cs:17-29 already made. FIX B moves the cell
boundary from `expression.Evaluate(...)` to `NamedReferences.CaptureValue(expression, context)` followed by
one new `ImplicitIntersection.Apply` arm in `Workbook.EvaluateCell` — capturing first is what makes bare
`=A1:A3` reachable at all (today `RangeReference.Evaluate` at RangeReference.cs:14 returns `#VALUE!` BEFORE
the boundary sees it, so a value-only arm would be a no-op), and `EvaluateCell` is provably the only choke
point (all of `GetCellValue`:199, `GetCellValueDense`:283, `GetCellValueOverflow`:311, `SheetValueReader`,
`RangeValueCache`:213/:287 funnel through it; warm-start restores only non-Reference surrogates). FIX C adds a
third backing to `PositionalRange` (the mini-CSE `ArrayStream`) reached through a new opt-in
`OpenArrayOrRange` factory that SUMPRODUCT alone uses, so the *IFS family keeps Excel's "arrays are #VALUE!
there" rule; the array backing is TRANSPOSED on read because `RangeValueSequence` is column-major while
`ArrayStream` is row-major, and without it a 2-D `SUMPRODUCT(A1:B2,(A1:B2)*1)` returns 29 instead of 30. I
built all of this against a copy of the tree and both suites stayed green (1203 + 88, zero failures), so the
honest answer to "what breaks" is: no existing test, because every current bare-range assertion
(ExpressionParserTests.cs:199) drives `Expression.Evaluate` directly rather than a cell.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** FIX C item 16: "extend the up-front validation at :35 to `if (ranges[a].Count != length || !SameShape(ranges[0], ranges[a]))`" with `SameShape` returning true whenever either side's `Rows == 0`. Presented as making SUMPRODUCT obey Excel's documented dimension rule.
      *Measured evidence:* The rule is silently bypassed for any range served from the per-epoch
      RangeSnapshot, because that backing goes through the LIST ctor which leaves Rows/Columns at 0 (patch:
      `private PositionalRange(IReadOnlyList<ComputedValue> list) { ... Rows = 0; Columns = 0; }`, over
      Danfma.MySheet/Expressions/CriteriaScan.cs:22-28), and the snapshot branch at CriteriaScan.cs:67-70 uses
      it. Admission is stateful second-use over 256 populated cells
      (Danfma.MySheet/RangeValueCache.cs:767-798, `RangeCacheMinimumCells = 256` at :723). MEASURED against
      the prototype at /tmp/mysheet-probe: three cells in ONE workbook holding the IDENTICAL formula
      `=SUMPRODUCT(A1:A300,A5:KN5)` (300x1 vs 1x300, 300 populated cells each) return ZZ1 = #VALUE! (cold),
      ZZ2 = 300, ZZ3 = 300. The unpatched repo returns 300 for all three. So the phase makes an Excel-
      documented rule depend on cache-admission order — two identical formulas disagree, and the same formula
      flips answer between a cold workbook and a warm one. Every SUMPRODUCT test the spec proposes uses 3-cell
      fixtures, far below the 256-cell threshold, so the proposed verification cannot detect this.
      *Correction:* Give the snapshot branch its shape: add a `PositionalRange(IReadOnlyList<ComputedValue>
      list, int rows, int columns)` ctor and, in `Open(argument, context, snapshot)`, derive `rows`/`columns`
      from `rectangle.GetBounds()` when the argument (after the AnchoredRangeReference unwrap) is a
      `RangeReference`; keep 0/0 only for the genuinely non-rectangular ArgumentFlattening fallback. Then add
      a >=256-cell SUMPRODUCT shape-mismatch test that is read from TWO cells, so the snapshot path is
      exercised.
- [ ] **B2.** Item 18: add `ROW(INDIRECT("A2"))=2` and `ROW(INDIRECT("A2:A3"))=2` to tests/Danfma.MySheet.Tests/Parsing/ReferenceFunctionTests.cs "using its existing `Grid`/`N`/`T` helpers (:8-25)", claiming "Every one of these values was measured against the patched engine".
      *Measured evidence:* ReferenceFunctionTests.cs:8-25 provides only `Grid`/`N`/`T` — there is no context-
      carrying `Calc`; the file's one existing ROW test evaluates with `ExpressionParser.Parse("=A5",
      sheet).Evaluate(workbook)` (:34-36), i.e. an EvaluationContext with an EMPTY SheetName.
      Indirect.TryResolveReference bails without a sheet name
      (Danfma.MySheet/Expressions/Lookup/Indirect.cs:50-56), so ReferencePosition.TryResolve's failure path
      re-evaluates the argument and propagates Indirect's own #REF!. MEASURED on the prototype, same formula,
      three harnesses: `Parse(f,sheet).Evaluate(workbook)` -> #REF!; `Parse(f,sheet).Evaluate(new
      EvaluationContext(wb, sheet.Name))` -> 2; stored in a cell and read via `GetCellValue` -> 2. The two
      INDIRECT rows of the proposed table therefore FAIL as written. (Every other row of the table reproduced
      exactly, in all three harnesses.)
      *Correction:* State the harness explicitly in the item: either evaluate `new EvaluationContext(workbook,
      sheet.Name)` (the idiom IndirectTests.cs:23-27 already uses) or store the formula in a cell and read it
      through `workbook.GetCellValue`. Also record in the spec that ReferencePosition's error-propagation
      changes the direct-path result for `ROW(INDIRECT(...))` from #VALUE! to #REF!.

## Major corrections

- [ ] **M1.** Item 10's rationale: "CaptureValue's closed list also carries AnchoredRangeReference, so a shared-formula master whose whole body is an anchored range gets the same treatment" — and risk 6's "the intersection is the SLAVE's own row because EvaluationContext.WithDelta leaves CellId untouched (SharedFormulaSlave.cs:16)".
      *Evidence:* Both halves are wrong. A shared-formula MASTER's expression is never an anchored node:
      WorksheetStreamLoader.cs:489 parses it with `ExpressionParser.ParseFormulaBody(formulaText,
      context.Sheet)`, so it is a plain RangeReference (already covered by CaptureValue's first arm). The
      SLAVES are `SharedFormulaSlave` wrappers (WorksheetStreamLoader.cs:230), and
      NamedReferences.CaptureValue (:59-69) has no SharedFormulaSlave arm — it falls to `_ =>
      expression.Evaluate(context)`, which reaches AnchoredRangeReference.Evaluate's unconditional #VALUE!
      (AnchoredRangeReference.cs:32-33). MEASURED on the prototype: Sheet1!C1 holding a bare
      `AnchoredRangeReference(A1:A3)` -> 5 (intersected); Sheet1!C3 holding `SharedFormulaSlave(that same
      node, +2, 0)` -> #VALUE!. A slave over a NameReference is fine (-> 9), so the gap is exactly "shared
      group whose body is a bare range": the master intersects, every slave stays #VALUE!.
      *Correction:* Add a `SharedFormulaSlave slave => CaptureValue(slave.Master,
      context.WithDelta(slave.DeltaRow, slave.DeltaColumn))` arm to NamedReferences.CaptureValue (it mirrors
      SharedFormulaSlave.TryResolveReference at SharedFormulaSlave.cs:33-34), and correct the two rationale
      sentences. Add a slave case to the new CellBoundaryIntersectionTests file.
- [ ] **M2.** Risk 3: "ROW/COLUMN in an ARRAY position covers only `[NameReference or Reference]`. A reference-returning FUNCTION argument stays scalar ... Extending the arm to Function would make ArrayEvaluation.Probe resolve — i.e. EVALUATE — that function's arguments, then TryBuildOperand resolve them a second time ... Do not widen the pattern without first adding a per-evaluation resolution memo."
      *Evidence:* The pattern is ALREADY wide enough to hit that exact defect, because `DynamicRange :
      Reference` (DynamicRange.cs:11) and its endpoints are arbitrary reference-returning functions:
      ResolveRowRange -> NamedReferences.TryResolveReference -> DynamicRange.TryResolveReference ->
      Index.TryResolveReference, which evaluates `Arguments[1]`/`Arguments[2]`
      (Danfma.MySheet/Expressions/Lookup/Index.cs:185, :194). Probe calls ResolveRowRange, then
      TryBuildOperand calls it again. MEASURED with a counting custom function:
      `=SUM(ROW(INDEX(A1:A3,CNT(),1):A3))` invokes `CNT()` TWICE on the prototype and ONCE on the unpatched
      repo (result 6 vs #VALUE!). This breaks two documented guarantees in the same file the spec promises to
      keep honest: ArrayEvaluation.cs:88-90 ("Scalar sub-expressions are still evaluated ONCE at build time")
      and :113-114 ("the subsequent build is guaranteed to succeed and is the SINGLE evaluation of the
      argument"), plus the :120-122 invariant `IsArrayEligible == (build result)` — with a volatile endpoint
      (`RANDBETWEEN`) Probe can resolve to a range (Array) while the build resolves to nothing (Scalar), after
      which `TryEvaluateStream` returns false and the consumer evaluates the argument a THIRD time on its
      scalar path.
      *Correction:* Either exclude `DynamicRange` from the new arm (`[NameReference or CellReference or
      RangeReference or OpenRangeReference or AnchoredRangeReference]`), or add the per-evaluation resolution
      memo the risk already prescribes, and rewrite risk 3 to say the double-resolve is live today for
      DynamicRange rather than hypothetical for Function. Add `SUM(ROW(INDEX(..,1,1):A3))` to the MiniCse test
      item so the shape is covered at all.

## Implementation items

- [ ] **1.** Create Danfma.MySheet/Expressions/Lookup/ReferencePosition.cs — `internal static class ReferencePosition` with `Row(Expression, EvaluationContext)`, `Column(Expression, EvaluationContext)` and a private `TryResolve(Expression, EvaluationContext, out Reference?, out ComputedValue failure)`. TryResolve calls `NamedReferences.TryResolveReference(argument, context, out reference, boundOpenRanges: false)`; on failure it evaluates the argument once and returns that value's error if it has one (else Error.Value); on success it re-runs `ReferenceGuard.MissingSheet(reference, context)` and fails with #REF! if the resolved target's sheet is gone. Row maps `CellReference => CellAddress.Parse(cell.Id).Row`, `RangeReference => range.TopRow`, `OpenRangeReference => open.RowMin ?? 1`, `_ => Error.Value`; Column is the mirror (`.Column`, `range.LeftColumn`, `open.ColMin ?? 1`).
      *Files:* `Danfma.MySheet/Expressions/Lookup/ReferencePosition.cs`
      *Why:* A separate file, not a `file static class`: OrderStatistics.cs:448 (`file static class
      OrderSelection`) and SumProducts.cs:93 (`SumOfPairs`) are invisible outside their file, and ROW lives in
      Row.cs while COLUMN lives in LookupFunctions.cs:272. boundOpenRanges:false copies the choice
      Rows.cs:17-29 documents in its own header comment — ToBoundedRange (OpenRangeReference.cs:346) returns
      the POPULATED box, so `boundOpenRanges:true` would make ROW(A:A) the first populated row instead of
      Excel's 1. The MissingSheet re-check on the RESOLVED reference is the pattern ReferenceGuard.cs:70-80
      already uses for NameReference and Subtotal.cs:181-186 uses after its own re-dispatch; without it
      `ROW(INDEX(Ghost!A1:A3,2,1))` would report a row number for a deleted sheet (ReferenceGuard.cs:96-98
      `default: return null` does not inspect a Function argument).
- [ ] **2.** In Danfma.MySheet/Expressions/Lookup/Row.cs, insert `[var only] => ReferencePosition.Row(only, context),` between the `[] when context.CellId is { } id` arm at :33 and `_ => ComputedValue.Error(Error.Value)` at :34. Keep all four existing syntactic arms (:15-16, :22-28) unchanged.
      *Files:* `Danfma.MySheet/Expressions/Lookup/Row.cs`
      *Why:* Row.cs:34's `_ => Error.Value` is exactly why ROW(INDEX(..)), ROW(name), ROW(INDIRECT(..)) and
      ROW(OFFSET(..)) all measure as #VALUE! today while ROWS (Rows.cs:17-29) and ISREF
      (InformationFunctions.cs:103) work. The existing arms stay as the zero-resolution fast path AND as the
      shared-formula delta path (AnchoredCellReference.Effective(context) /
      AnchoredRangeReference.ToRangeReference(context) at :22-28) — a generic resolve would reach the same
      answer but pay a virtual call and lose the comment trail. `[var only]` must come AFTER the `[]` arm or a
      zero-argument ROW() would not match.
- [ ] **3.** In Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs, insert `[var only] => ReferencePosition.Column(only, context),` in `record Column` (:272) between the `[] when context.CellId is { } id` arm at :293-295 and `_ => ComputedValue.Error(Error.Value)` at :296.
      *Files:* `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs`
      *Why:* COLUMN is the byte-for-byte mirror of Row.cs (its own comments at :284-286 say so). Measured
      today: COLUMN(INDEX(A1:C1,1,2)) = #VALUE!, COLUMN(MyName) = #VALUE!, COLUMN(A:A) = #VALUE!; after the
      arm they are 2, 1, 1 — Excel's answers.
- [ ] **4.** Change `ArrayEvaluation.IsArrayEligible(Expression)` (ArrayEvaluation.cs:118) to `IsArrayEligible(Expression, EvaluationContext)` and `Probe(Expression)` (:123) to `Probe(Expression, EvaluationContext)`, threading the context through the five recursive Probe calls (binary Left/Right, If Arguments[0]/[1]/[2]) and updating the four call sites: NumericAggregation.cs:110, Index.cs:24, Index.cs:171, OrderStatistics.cs:508.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`, `Danfma.MySheet/Expressions/NumericAggregation.cs`, `Danfma.MySheet/Expressions/Lookup/Index.cs`, `Danfma.MySheet/Expressions/Statistical/OrderStatistics.cs`
      *Why:* Unavoidable, and this is the crux of the array half of FIX A: whether ROW(x) is an ARRAY or a
      SCALAR depends on the resolved reference's SHAPE (multi-cell range vs single cell), which no purely
      syntactic predicate can know for a NameReference or a table-column node. Probe's own header (:120-122)
      requires it to 'track the builder's structure exactly so IsArrayEligible == (build result)', and
      TryBuildOperand (:420) already takes a context — so the only way to keep that invariant is to give Probe
      the same information. No test calls IsArrayEligible (tests only drive TryEvaluate, which already takes a
      context: ArrayEvaluationTests.cs:53-270), so the blast radius is those four lines.
- [ ] **5.** Add to ArrayEvaluation.cs, just above `BuildRange` (:483): `private enum RowArgumentShape { Array, Scalar, Refused }` and `private static RowArgumentShape ResolveRowRange(Expression argument, EvaluationContext context, out RangeBounds bounds)` — resolve via `NamedReferences.TryResolveReference(argument, context, out var reference, boundOpenRanges: false)`; a `RangeReference` sets `bounds = range.GetBounds()` and returns Array, an `OpenRangeReference` returns Refused, everything else (a single cell, an unresolved name) returns Scalar.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* One resolver shared by Probe and TryBuildOperand is what keeps the documented 'eligible iff the
      build succeeds' invariant (:110-116) intact — the alternative (Probe guessing, the build failing) would
      break the 'the subsequent build is the SINGLE evaluation' guarantee. boundOpenRanges:false is required
      so an OpenRangeReference reaches the Refused arm instead of being silently collapsed to a populated box,
      preserving the cost guard the file states at :44-46 and :465-468. RangeBounds is internal in the same
      namespace (RangeReference.cs:170-180) and GetBounds() must be called once (its own doc warns each call
      parses both corners).
- [ ] **6.** Add to `ArrayEvaluation.Probe`, immediately AFTER the `case Row { Arguments: [OpenRangeReference] }` refusal at :149: `case Row { Arguments: [NameReference or Reference] } resolvable:` returning `(true,true)` / `(false,false)` / `(true,false)` from `ResolveRowRange(resolvable.Arguments[0], context, out _)` for Array/Refused/Scalar. Add the mirror arm to `TryBuildOperand` immediately after its `case Row { Arguments: [OpenRangeReference] }` at :465-468, building `new RowNumbersOperand(bounds.TopRow, bounds.RowCount, bounds.ColumnCount)` on Array, returning false on Refused, and `new ScalarOperand(expression.Evaluate(context))` on Scalar.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* This is the edit that makes SUM(ROW(name)) = 6 instead of #VALUE! and COUNT(ROW(name)) = 3
      instead of 0 (both measured). Position matters: the arm MUST come after the three existing ROW arms at
      :143-151 / :459-468 (RangeReference, AnchoredRangeReference, OpenRangeReference) or
      `Row{[RangeReference]}` would be captured by `[... or Reference]` and pay a needless resolve. The `or
      Reference` half is the deliberate hook for Phase 3: a table-column node that derives from `Reference`
      and returns a concrete RangeReference from TryResolveReference (the shape the reference-model facts
      recommend) is covered by this arm with ZERO further edit; DynamicRange is covered too (measured:
      SUM(ROW(INDEX(A1:A3,1,1):A3)) = 6).
- [ ] **7.** Add `private sealed class ColumnNumbersOperand : ArrayOperand` to ArrayEvaluation.cs above `BinaryOperand` (:330), mirroring `RowNumbersOperand` (:302-328) but returning `ComputedValue.Number(_leftColumn + index % _columns)`; then add four `Column {...}` arms to Probe (RangeReference / AnchoredRangeReference => (true,true), OpenRangeReference => (false,false), `[NameReference or Reference]` => ResolveRowRange) and the four mirror arms to TryBuildOperand, using `bounds.LeftColumn`.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* Measured today: SUM(COLUMN(A1:C1)) = 1 where Excel gives 6, because RowNumbersOperand has no
      COLUMN twin — leaving ROW arrayed and COLUMN not would be a fresh asymmetry inside the same fix.
      Verified after the change: SUM(COLUMN(A1:C1))=6, SUM(COLUMN(A1:C3))=18, SUM(COLUMN(MyName))=3,
      SMALL(COLUMN(A1:C1),2)=2, and SUMPRODUCT(ROW(A1:B2),COLUMN(A1:B2))=9 (which also cross-checks that both
      operands agree on row-major order). This item is separable: dropping it costs no corpus formula, since
      the corpus uses ROW.
- [ ] **8.** Update the ArrayEvaluation.cs doc comments to match: the class summary's supported set (:36-46) gains 'ROW/COLUMN over a name or a structured reference', `IsArrayEligible`'s remark (:104-116) records that it now takes a context because a resolvable ROW/COLUMN argument's SHAPE is context-dependent, and Probe's invariant comment (:120-122) names ResolveRowRange as the shared oracle.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* The file's contract is load-bearing documentation for four consumers; :110-116 currently promises
      a 'CHEAP syntactic pre-check — no evaluation', which is no longer literally true for a ROW(name)
      argument (it costs one dictionary lookup plus a virtual call). Leaving that comment stale is exactly the
      kind of drift that produced the STALE tag-policy comment at Expression.cs:14-16.
- [ ] **9.** Create Danfma.MySheet/Expressions/ImplicitIntersection.cs — `internal static class ImplicitIntersection` with `Apply(Reference reference, EvaluationContext context)`. It reads the formula cell's own column/row from `context.CellId` via `CellAddress.TryGetColumnRow` (no A1 id → Error.Value), then: `CellReference` → dereference; `RangeReference` → `range.GetBounds()` into a private `Intersect(context, sheetName, int? left, int? right, int? top, int? bottom, formulaColumn, formulaRow)`; `OpenRangeReference` → the same Intersect with the DECLARED `ColMin/ColMax/RowMin/RowMax` (nulls = unbounded); `default` → Error.Value. Intersect: 1x1 → that cell; single column and formulaRow within [top,bottom] → (left, formulaRow); single row and formulaColumn within [left,right] → (formulaColumn, top); otherwise Error.Value. Dereference through `workbook.GetCellValueDense(workbook.ResolveDenseHandle(sheetName), sheetName, column, row)`.
      *Files:* `Danfma.MySheet/Expressions/ImplicitIntersection.cs`
      *Why:* DECLARED bounds, not populated: `=A:A` in C7 must be A7 (measured: 0, because A7 is empty and the
      boundary's blank→0 rule applies) — `OpenRangeReference.ToBoundedRange` (:346) returns the POPULATED box
      and would give A1..A3's intersection instead, so it must NOT be used here. Nullable bounds let one
      Intersect serve both closed and open shapes: `A:A` is single-column with an unbounded row axis, `1:1`
      single-row with an unbounded column axis; measured `=1:1` in C7 = C1's value, `=A2:A` in C1 = #VALUE!
      and in C3 = A3, `=A:A10` in C77 = #VALUE!. GetCellValueDense is the allocation-free deref idiom already
      used by RangeReference.CellComputedValueAt (RangeReference.cs:160). Cross-sheet is positional and sheet-
      independent by construction: measured `=Sheet1!A1:A3` in Sheet2!C2 = Sheet1!A2 = 0, and `=Ghost!A1:A3`
      in C2 = #REF! because the deref lands in EvaluateCell's missing-sheet arm (Workbook.cs:339-342) — no
      extra guard needed.
- [ ] **10.** Patch `Workbook.EvaluateCell` (Danfma.MySheet/Workbook.cs:346-347): hoist `var context = new EvaluationContext(this, sheetName, id);`, replace `value = expression.Evaluate(context)` with `value = NamedReferences.CaptureValue(expression, context);`, then add `if (value.TryGetReference(out var boundaryReference)) { value = ImplicitIntersection.Apply(boundaryReference, context); }`. Leave the `expression is not BlankValue && value.Kind == Blank` coercion at :355-357 BELOW the new block, unchanged.
      *Files:* `Danfma.MySheet/Workbook.cs`, `Danfma.MySheet/Expressions/NamedReferences.cs`
      *Why:* CaptureValue (NamedReferences.cs:59-69) is the exact primitive that already means 'a top-level
      range denotes a reference value, not #VALUE!' — reusing it instead of inventing a boundary type-switch
      means the boundary and a defined-name/LET binding agree by construction, and it is what makes bare
      `=A1:A3` reachable (RangeReference.Evaluate at RangeReference.cs:14 fires first otherwise, and its
      #VALUE! is indistinguishable from a real one). Order is load-bearing: capture → intersect → blank→0, so
      `=A:A` in C7 with A7 empty yields 0 like Excel (verified) while a truly BlankValue cell still stays
      blank (verified). CaptureValue's closed list also carries AnchoredRangeReference, so a shared-formula
      master whose whole body is an anchored range gets the same treatment. Measured per-cell correctness:
      `=A1:A4` gives 10 in C1 and 30 in C3 — the intersection is the SLAVE's own row because
      EvaluationContext.WithDelta leaves CellId untouched (SharedFormulaSlave.cs:16).
- [ ] **11.** Update the `<para>` in CachedCellValue.cs:11-14 and the comment on the `case ComputedValueKind.Reference` arm at :64-67 to say the arm is now defence-in-depth over the PUBLIC ComputedValue type (a host custom function can still hand back `ComputedValue.Reference`, ComputedValue.cs:51) rather than a live exclusion, since the store can no longer hold a Reference. Do NOT delete the arm.
      *Files:* `Danfma.MySheet/CachedCellValue.cs`
      *Why:* Workbook.Serialization.cs:377 feeds TryFrom straight from `store.EnumerateNonTainted()`, and
      every store write is one of Workbook.cs:200/:283/:312 — all immediately after EvaluateCell — so after
      the boundary arm no stored value can be Reference and the exclusion becomes unreachable via that path.
      It must survive anyway: TryFrom takes a public ComputedValue, and WarmStartSaveLoadTests.cs:384 asserts
      `TryFrom(..., reference) is null` by constructing the value directly. Verified the warm path round-trips
      the new value: `=MyName` in C2 saved and reloaded as 0 (it was previously dropped from the snapshot and
      recomputed).
- [ ] **12.** Add a one-line addendum to the `NameReference` arm comment in Danfma.MySheet/Parsing/AnchoredFormulaSupport.cs:37 noting that 'position-independent' means the RESOLVED reference is delta-invariant — the cell-boundary implicit intersection is applied per slave cell from its own CellId, so a shared group over `=SomeName` still yields a different value per row and remains anchored-safe.
      *Files:* `Danfma.MySheet/Parsing/AnchoredFormulaSupport.cs`
      *Why:* IsFullyAnchored is consulted at exactly one place, WorksheetStreamLoader.cs:267, purely to choose
      the anchored-vs-legacy PARSE strategy for a shared-formula group — no computed value is shared across
      the group, so the verdict stays correct. But :37's bare claim 'a defined name is position-independent'
      now reads as false at a glance, and a future reader could rip the arm out. Comment-only; verified by the
      shared-formula probe (C1=10, C3=30 for the same formula text).
- [ ] **13.** Update the comments on the now-unreachable `ComputedValueKind.Reference` arms in Danfma.MySheet.Excel/ExcelExport.cs:279 and :292 and Danfma.MySheet.Excel/ExcelMerge.cs:575-579 to say a cell value can no longer BE a reference (the boundary intersects it) and that the arm survives only as a total-switch guard. Keep all three arms.
      *Files:* `Danfma.MySheet.Excel/ExcelExport.cs`, `Danfma.MySheet.Excel/ExcelMerge.cs`
      *Why:* Both read through `workbook.GetCellValue(...)` (ExcelExport.cs:180; ExcelMerge.cs:419, :472,
      :523), so they are downstream of EvaluateCell and can never see Reference again. This is a USER-VISIBLE
      export change to call out in the docs: a cell holding `=MyName` used to export as t="e" #VALUE!
      (documented at docs/excel-interop.md:161-162) and now exports the intersected value — which is what
      Excel itself would have stored, so it is an improvement, but it is a change.
- [ ] **14.** In Danfma.MySheet/Expressions/CriteriaScan.cs, give `PositionalRange` a third backing: fields `private readonly ArrayEvaluation.ArrayStream _stream; private readonly int _streamRows; private readonly int _streamColumns;`, a `private PositionalRange(ArrayEvaluation.ArrayStream stream)` ctor setting `Count = stream.Length`, and public readonly `Rows`/`Columns` shape fields (0 = not rectangular) set from the bounds in the cursor ctor (change its signature from `(cursor, int count)` at :30 to `(cursor, int rows, int columns)` and pass `bounds.RowCount, bounds.ColumnCount` at the call site inside Open) and from `stream.Rows/Columns` in the new ctor; 0/0 for the list ctor at :22.
      *Files:* `Danfma.MySheet/Expressions/CriteriaScan.cs`
      *Why:* PositionalRange already documents 'Exactly one backing is live' (:12-13) and already carries an
      up-front `Count` so the *IFS validation never materializes to measure (:17-19) — a third backing plus a
      shape is the smallest change that keeps that discipline while letting SUMPRODUCT read a computed vector
      without a `ComputedValue[]`. The shape is needed by the SameShape check below: Count alone cannot tell a
      3x1 column from a 1x3 row.
- [ ] **15.** Add `public static PositionalRange OpenArrayOrRange(Expression argument, EvaluationContext context)` to PositionalRange: when `argument is not Reference && ArrayEvaluation.IsArrayEligible(argument, context) && ArrayEvaluation.TryEvaluateStream(argument, context, out var stream)` return the stream backing, else delegate to the existing `Open(argument, context)`. Document that it is SUMPRODUCT-only.
      *Files:* `Danfma.MySheet/Expressions/CriteriaScan.cs`
      *Why:* An opt-in factory rather than folding the gate into `Open` (:42-48), because Open is shared with
      CriteriaScan (SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS) where Excel REQUIRES real ranges —
      `SUMIFS((A1:A3)*1, …)` is #VALUE! in Excel, so silently enabling arrays there would be a semantics
      change outside this scope. The three-condition gate is copied verbatim from OrderStatistics.cs:506-513
      and NumericAggregation.cs:110-111, so SUMPRODUCT inherits the same 'single evaluation, no double-eval'
      property.
- [ ] **16.** In `PositionalRange.Next()` (CriteriaScan.cs:99-108), insert the array branch between the list branch and the cursor branch: `if (_streamRows > 0) { var position = _index++; var column = position / _streamRows; var row = position % _streamRows; return _stream.ElementAt(row * _streamColumns + column); }`.
      *Files:* `Danfma.MySheet/Expressions/CriteriaScan.cs`
      *Why:* This transpose is load-bearing, not defensive. RangeValueSequence.Enumerator walks COLUMN-major
      (RangeReference.cs:186-190 header: 'column-major (row-inner) order') while ArrayStream.ElementAt is ROW-
      major (ArrayEvaluation.cs:591-600). Measured with A1=1,A2=2,B1=3,B2=4: `SUMPRODUCT(A1:B2,(A1:B2)*1)` =
      30 with the transpose and 29 without — a silently wrong number for every 2-D mix of a range and a
      computed array. `_streamRows > 0` doubles as the backing discriminator, so no extra bool field.
- [ ] **17.** In Danfma.MySheet/Expressions/Mathematics/SumProducts.cs, switch both `PositionalRange.Open` calls (:26 and :33) to `PositionalRange.OpenArrayOrRange`, and extend the up-front validation at :35 to `if (ranges[a].Count != length || !SameShape(ranges[0], ranges[a]))`. Add `private static bool SameShape(in PositionalRange first, in PositionalRange other) => first.Rows == 0 || other.Rows == 0 || (first.Rows == other.Rows && first.Columns == other.Columns);` as a member of the SumProduct record (NOT in `file static class SumOfPairs`).
      *Files:* `Danfma.MySheet/Expressions/Mathematics/SumProducts.cs`
      *Why:* Two changes in one edit because they are inseparable: enabling arrays makes
      `SUMPRODUCT((A1:A3<>0)*1, A1:C1)` reachable, and with a count-only check it returns 5 where Excel
      returns #VALUE! — the rule is quoted verbatim in the repo's own test at MathAggregateTests.cs:83-84
      ('The array arguments must have the same dimensions. If they do not, SUMPRODUCT returns the #VALUE!
      error value'), so shape validation is required by P0, not optional. It also fixes the pre-existing
      range-vs-range case: `SUMPRODUCT(A1:A3,A1:C1)` measured 25 before and #VALUE! after. `SameShape` cannot
      live in `file static class SumOfPairs` (:93) — that type is file-scoped and the record needs it; and it
      must tolerate Rows==0, or `SUMPRODUCT(MyName, (A1:A3<>0)*1)` would be rejected (a NameReference takes
      the materialized fallback whose shape is unknown; measured 14 both before and after). Everything else in
      the body is untouched, so errors still propagate at :53-56 and non-numerics still zero at :59 — both
      Excel-correct.
- [ ] **18.** Add a ROW/COLUMN resolution table to tests/Danfma.MySheet.Tests/Parsing/ReferenceFunctionTests.cs using its existing `Grid`/`N`/`T` helpers (:8-25), on the fixture A1=5,A2=0,A3=9 with `MyName = Sheet1!$A$1:$A$3`, `MyCell = Sheet1!$A$2`: ROW(A1:A3)=1, ROW(A2)=2, ROW(INDEX(A1:A3,2,1))=2, ROW(OFFSET(A1,1,0))=2, ROW(OFFSET(A1,1,0,2,1))=2, ROW(INDIRECT("A2"))=2, ROW(INDIRECT("A2:A3"))=2, ROW(MyName)=1, ROW(MyCell)=2, ROW(CHOOSE(1,A2:A3,B1))=2, ROW(A:A)=1, ROW(A2:A)=2, ROW(INDEX(A1:A3,2,1):A3)=2, ROW(NoSuchName)=#NAME?, ROW(Ghost!A1:A3)=#REF!, ROW(1)=#VALUE!, ROW((A1:A3,B1:B3))=#VALUE!; and COLUMN(A1:C1)=1, COLUMN(INDEX(A1:C1,1,2))=2, COLUMN(MyName)=1, COLUMN(A:A)=1, COLUMN(1:1)=1.
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/ReferenceFunctionTests.cs`
      *Why:* Every one of these values was measured against the patched engine, and each is Excel's documented
      answer for ROW/COLUMN of a reference (top row / leftmost column of the resolved reference). The #NAME?
      and #REF! cases pin the two failure arms of ReferencePosition.TryResolve, which are the arms most likely
      to be 'simplified' away later; ROW((A1:A3,B1:B3))=#VALUE! pins the deliberate no-change for unions.
      ReferenceFunctionTests.cs currently holds exactly one ROW test (:28-36, ROW() with no argument) — this
      is the natural home.
- [ ] **19.** Add array-path cases to tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs (its own fixture helper style, plus a numeric grid): SUM(ROW(MyName))=6, COUNT(ROW(MyName))=3, SMALL(ROW(MyName),1)=1, INDEX(ROW(MyName),2)=2, SUM(ROW(A1:C3))=18, SUM(COLUMN(A1:C1))=6, SUM(COLUMN(A1:C3))=18, SUM(COLUMN(MyName))=3, SMALL(COLUMN(A1:C1),2)=2, SUMPRODUCT(ROW(A1:B2),COLUMN(A1:B2))=9, plus the two cost-guard negatives SUM(ROW(A:A))=1 and SUM(COLUMN(A:A))=1.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`
      *Why:* MiniCseConsumerTests.cs:7-12 declares itself the home of the mini-CSE production consumers, and
      COUNT(ROW(MyName)) is the sharpest regression sentinel: it measures 0 today (the fold treats the scalar
      row number as a single element that the *A-less COUNT then... counts as one — the exact silent-wrong-
      answer class this fix removes) and 3 after. The cost-guard negatives pin ResolveRowRange's Refused arm,
      without which an open column would materialize.
- [ ] **20.** Create tests/Danfma.MySheet.Tests/Expressions/CellBoundaryIntersectionTests.cs — put a formula in a CELL and read it via `workbook.GetCellValue(sheet, id)` (NOT `Expression.Evaluate`), fixture A1=5,A2=0,A3=9,B1=2,D1="txt",E1==1/0, Sheet2!A1..A3 = 11,22,33, MyName=Sheet1!$A$1:$A$3, Ghosty=Ghost!$A$1:$A$3. Assert: `=A1:A3` in C2 → 0, in C9 → #VALUE!, in A2 → #REF!; `=A1:C1` in B5 → 2; `=A1:C3` in B2 and C1 → #VALUE!; `=A1:A1` in Z99 → 5; `=A:A` in C7 → 0; `=1:1` in C7 → 0; `=A2:A` in C1 → #VALUE! and in C3 → 9; `=A:A10` in C77 → #VALUE!; `=MyName` in C2 → 0; `=Sheet1!A1:A3` in Sheet2!C2 → 0; `=Ghost!A1:A3` and `=Ghosty` in C2 → #REF!; `=(A1:A3,B1:B3)` in C2 → #VALUE!; `=D1:D3` in C1 → "txt"; `=E1:E3` in C1 → #DIV/0!; `=INDIRECT("MyName")`, `=INDIRECT("A1:A3")`, `=OFFSET(A1,0,0,3,1)`, `=CHOOSE(1,A1:A3)`, `=+A1:A3`, `=LET(x,A1:A3,x)`, `=INDEX(A1:A3,2,1):A3` in C2 → 0; and the unchanged controls `=SUM(A1:A3)` in C2 → 14, `=A1` in C2 → 5, an empty C5 → blank. Also assert `ExpressionParser.Parse("=A1:A3", sheet).Evaluate(workbook)` is STILL #VALUE! (the direct path is not a cell boundary), and that a saved+reloaded workbook returns the same intersected value for `=MyName` in C2.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/CellBoundaryIntersectionTests.cs`
      *Why:* Every listed value is measured against the patched engine. A NEW file rather than extending
      IndirectTests.cs, because the distinguishing feature of all these cases is the READ PATH:
      IndirectTests.cs:23-27's own `Calc` helper calls `Expression.Evaluate` directly and therefore never
      crosses the boundary — which is exactly why no existing test broke and exactly the trap the next reader
      will fall into. The `=A1:A3` in A2 → #REF! case documents that self-intersection hits the cycle guard
      (Workbook.cs:190-193) where Excel shows a circular-reference dialog. The direct-path assertion is the
      honest boundary of the change and stops someone 'fixing' RangeReference.Evaluate (RangeReference.cs:14)
      instead.
- [ ] **21.** Add SUMPRODUCT cases to tests/Danfma.MySheet.Tests/Parsing/MathAggregateTests.cs under its existing golden header (:6-8, support.microsoft.com fetched 2026-07-02, SUMPRODUCT page 16753e75-9f68-4874-94ac-4d2145a2fd2e) using the existing `Calc`/`Num` helpers: on A1=5,A2=0,A3=9,B1=2,B2=4,B3=6 — SUMPRODUCT((A1:A3<>0)*1)=2, SUMPRODUCT((A1:A3<>0)*1,B1:B3)=8, SUMPRODUCT((A1:A3>0)*(B1:B3>3))=1, SUMPRODUCT(ROW(A1:A3))=6, SUMPRODUCT(IF(A1:A3>0,1,0))=2, SUMPRODUCT((A1:A3<>0))=0 (logicals are not coerced), SUMPRODUCT(A1:A3,A1:C1)=#VALUE! and SUMPRODUCT((A1:A3<>0)*1,A1:D1)=#VALUE! (dimension rule), SUMPRODUCT((E1:E3)*1,A1:A3)=#DIV/0! with E1==1/0; and on A1=1,A2=2,B1=3,B2=4 the transpose quartet SUMPRODUCT(A1:B2,A1:B2)=30, SUMPRODUCT(A1:B2,(A1:B2)*1)=30, SUMPRODUCT((A1:B2)*1,A1:B2)=30, SUMPRODUCT((A1:B2)*1,(A1:B2)*1)=30.
      *Files:* `tests/Danfma.MySheet.Tests/Parsing/MathAggregateTests.cs`
      *Why:* SUMPRODUCT(ROW(A1:A3)) is the most important of these: it returns 1 today — a WRONG NUMBER, not
      an error — because PositionalRange.Open falls through to ArgumentFlattening and evaluates ROW(range) as
      one scalar. The transpose quartet is the only test that would catch a regression of the column-
      major/row-major reconciliation (all four must be 30; drop the transpose and the two mixed forms become
      29). The dimension cases sit next to the existing SumProduct_ShapeMismatch_IsValueError (:78-86) which
      already quotes the same Microsoft sentence, so the golden-citation convention is satisfied by adjacency.
      No FormulaWriterTests.cs change is needed: this phase adds no AST node, and parse→write identity was
      verified for =A1:A3, =ROW(MyName), =COLUMN(A1:C1), =SUMPRODUCT((A1:A3<>0)*1), =A:A and
      =ROW(INDIRECT("A2")).
- [ ] **22.** Rewrite docs/computed-value.md:166-170 (and the pt-BR mirror docs/pt-BR/computed-value.md:169-173): a bare range in a CELL no longer yields #VALUE! — it is implicitly intersected with the formula cell's row/column, and `ComputedValueKind.Reference` can no longer be a cell's cached value (it still occurs from `Expression.Evaluate` on a reference-returning function, which is what EnumerateValues is for). Update the §References intro at :147-150 accordingly and keep the OFFSET/EnumerateValues example.
      *Files:* `docs/computed-value.md`, `docs/pt-BR/computed-value.md`
      *Why:* docs/computed-value.md:166-169 states the contract this phase reverses ('a bare range ...
      evaluates to #VALUE!, as in Excel'), and the parenthetical 'as in Excel' is the part that was wrong —
      Excel returns the intersected cell. docs/pt-BR/README.md:3 makes English authoritative, but the mirror
      is confirmed complete (11 files each) so both must move together.
- [ ] **23.** In docs/workbook-and-expressions.md, replace the bare-range bullet at :282-283 with a pointer to a NEW section '## Implicit intersection at the cell boundary' placed after §Whole-column and whole-row references, stating the rule (single-column range spanning the formula's row → that row's cell; single-row range spanning the formula's column → that column's cell; a 1x1 range → itself; no intersection or a 2-D range → #VALUE!; declared bounds for open ranges so `=A:A` in row 7 is A7; positional and sheet-independent so `=Sheet2!A1:A3` in Sheet1!C2 is Sheet2!A2), the explicit no-spill deviation, that it applies to `=A1:A3`, `=MyName`, `=INDIRECT("Name")`, `=OFFSET(..)`, `=CHOOSE(..)`, `=LET(x,range,x)` and structured table references alike, and that the DIRECT `Expression.Evaluate` path still yields #VALUE! (there is no formula cell to intersect with). Cross-reference it from the 'dry cell keeps #VALUE!' bullet at :361-363 (which stays TRUE — a computed IF-array is not a Reference) and from the §Named ranges area. Mirror into docs/pt-BR/workbook-and-expressions.md (bullet at :290-291, same new section).
      *Files:* `docs/workbook-and-expressions.md`, `docs/pt-BR/workbook-and-expressions.md`
      *Why:* This is the primary user-facing contract change and it needs its own section, not a patched
      bullet: the rule has five cases and two deliberate deviations (no spill; 2-D → #VALUE!). Keeping the
      :361-363 mini-CSE bullet and clarifying why it is unaffected matters — measured `=IF(TRUE,A1:A3,B1)` in
      a cell is still #VALUE! because If.Evaluate returns the branch's own Evaluate rather than routing
      through CaptureValue, so the two statements coexist but look contradictory without a cross-reference.
- [ ] **24.** Update the ROW and COLUMN rows of docs/function-reference.md (:245 and :237) and docs/pt-BR/function-reference.md (:251 and :243) to say the argument may be ANY reference-producing expression (a defined name, a structured table reference, INDIRECT/INDEX/OFFSET/CHOOSE, a dynamic ':' range), that a whole-column/row reference reports its declared lower bound (ROW(A:A)=1, COLUMN(1:1)=1) — contrast with the ROWS/COLUMNS populated-extent rule already documented on the adjacent rows — and that ROW/COLUMN of a name or structured reference also participates in the [implicit array arguments](workbook-and-expressions.md#implicit-array-arguments) path. Add ROW/COLUMN-over-a-name to the §Implicit array arguments supported list at docs/workbook-and-expressions.md:352-356 and the pt-BR twin. Add SUMPRODUCT's array acceptance to its function-reference row in both languages.
      *Files:* `docs/function-reference.md`, `docs/pt-BR/function-reference.md`, `docs/workbook-and-expressions.md`, `docs/pt-BR/workbook-and-expressions.md`
      *Why:* Function counts are untouched (no new built-in), so only prose changes — but §Implicit array
      arguments at :352-356 currently enumerates the eligible set as 'a closed-range comparison, an IF whose
      condition is such an array, or ROW(range)', which becomes incomplete. The ROW(A:A)=1 vs
      ROWS(A:A)=populated-extent contrast is worth spelling out because the two now use the SAME
      boundOpenRanges:false resolution but different rules on the open axis, and the pt-BR ROWS row (:252)
      already documents its divergence from Excel's fixed grid.
- [ ] **25.** Add one sentence to docs/custom-functions.md:128-129 and docs/pt-BR/custom-functions.md:137 clarifying that 'evaluating a range directly yields #VALUE!' describes `Expression.Evaluate`, and that a range reaching a CELL's value is instead implicitly intersected (link the new section); and add a bullet to docs/excel-interop.md near :161-162 recording that a cell whose value used to export as t="e" #VALUE! for a bare-reference result now exports the intersected value.
      *Files:* `docs/custom-functions.md`, `docs/pt-BR/custom-functions.md`, `docs/excel-interop.md`, `docs/pt-BR/excel-interop.md`
      *Why:* docs/custom-functions.md:128-129 stays literally true (Expression.Evaluate is unchanged) but now
      reads as a contradiction of the new section; one clause fixes it. docs/excel-interop.md:161-162
      documents the #VALUE! that ExcelExport.cs:279/:292 and ExcelMerge.cs:575-579 wrote, and that output
      genuinely changes for users who round-trip a workbook containing `=SomeName` — recording it is the
      difference between a documented improvement and a surprise.
- [ ] **26.** Commit FIX B as `feat(eval):` (a documented-behaviour change to the cell boundary), FIX A as `feat(eval):`, FIX C as `fix(eval):`; never hand-edit CHANGELOG.md — versionize generates it from the Conventional Commit subjects, and S7 fixes the release at 3.17.0.
      *Files:* `CHANGELOG.md`
      *Why:* CLAUDE.md/project discipline: CHANGELOG.md is versionize-generated and the scopes in use are
      parser/eval/serialization/recalc/dirty-graph. FIX B reverses a contract stated in docs/computed-
      value.md:166-169 and docs/workbook-and-expressions.md:282-283, so classifying it as `fix` would
      understate it in the release notes and could land it as a patch bump; FIX C repairs a wrong NUMBER
      (SUMPRODUCT(ROW(A1:A3)) = 1) so `fix` is right for it.

## Scope addition — decided 2026-09-08 under P0

`ROWS`, `COLUMNS` and `AREAS` propagate an unresolvable argument's error exactly as the new `ROW`/`COLUMN`
do, instead of answering `1`. Today `ROWS(NoSuchName)` = 1 (Rows.cs `: 1.0` fallback) where Excel gives
`#NAME?`, and `ROWS(INDIRECT("zz"))` = 1 where Excel gives `#REF!`. This is a behaviour change to three
shipped functions and ships as `feat(eval):`; it exists so that `ROW` and `ROWS` cannot disagree about the
same invalid argument (see risk "ReferencePosition's failure path…" below, which this resolves).

- [ ] **27.** Route the unresolvable-argument fallback of `Rows` (Rows.cs), `Columns` and `Areas` (both in
      LookupFunctions.cs) through the same `ReferencePosition` error-recovery helper FIX A adds, so the three
      return the argument's own error (`#NAME?` for an unknown name, `#REF!` for a failed INDIRECT/OFFSET)
      rather than `1`. Keep the successful paths untouched. TDD: pin `ROWS(NoSuchName)`, `COLUMNS(NoSuchName)`,
      `AREAS(NoSuchName)` → `#NAME?` and `ROWS(INDIRECT("zz"))` → `#REF!` in
      tests/Danfma.MySheet.Tests/Parsing/ReferenceFunctionTests.cs, evaluating with
      `new EvaluationContext(workbook, sheet.Name)` (B2 above explains why the bare `Evaluate(workbook)` harness
      cannot see INDIRECT).
      *Files:* `Danfma.MySheet/Expressions/Lookup/Rows.cs`, `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs`,
      `tests/Danfma.MySheet.Tests/Parsing/ReferenceFunctionTests.cs`, `docs/function-reference.md`,
      `docs/pt-BR/function-reference.md`
      *Why:* P0. Excel propagates the error for every one of these; the `1` fallback is a MySheet invention
      that returns a plausible number for a broken reference. Doing it in the same phase as FIX A means one
      helper, one error-precedence rule and one test file, instead of a documented divergence between `ROW` and
      `ROWS`.

## Verification Plan

- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet build Danfma.MySheet.slnx -c Release`
      → expected: "Build succeeded." with "0 Warning(s)" and "0 Error(s)". Verified on the prototype: the
      whole design (2 new engine files, 9 modified) compiles warning-free.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier check .`
      → expected: Ends with "Checked N files in …" and prints NO "Was not formatted" line; exit code 0. The
      prototype needed `dotnet csharpier format .` first for two long lines (Index.cs:171's && chain and the
      AnchoredRangeReference COLUMN arm) — run format, then check.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release`
      → expected: "Passed!" with "failed: 0". total must be 1203 + the number of new [Test] methods added by
      this phase. The 1203 pre-existing tests were all verified GREEN against the complete prototype — any
      pre-existing failure means the implementation diverged from this design, it is not expected fallout.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj -c Release`
      → expected: "Passed!" with "total: 88" and "failed: 0" — unchanged. Verified against the prototype,
      including the six TableInteropTests and the export/merge cached-value tests.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/CellBoundaryIntersectionTests/*"`
      → expected: "Passed!", failed: 0, total equal to the number of tests in the new file. Confirms the FIX B
      table specifically; the --treenode-filter form was verified working on this TUnit version
      (ReferenceFunctionTests → total: 11).
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/MathAggregateTests/*"`
      → expected: "Passed!", failed: 0. In particular the transpose quartet must ALL be 30: if
      SUMPRODUCT(A1:B2,(A1:B2)*1) reports 29, the column-major transpose in PositionalRange.Next() is missing
      or wrong.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/CellStoreTests/*" && dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release -- --treenode-filter "/*/*/WarmStartSaveLoadTests/*"`
      → expected: Both "Passed!", failed: 0 — proving this phase touches no wire format:
      PreChangeCellsWireGolden (CellStoreTests.cs:20/:61-66, the frozen "AgIAAAD7…" base64) and
      Save_Default_IsByteIdenticalToRawMemoryPack (WarmStartSaveLoadTests.cs:31-49) both passed against the
      prototype. No union tag and no new Workbook member is introduced here.
- [ ] `cd /Volumes/Work/Develop/MySheet && git diff --stat -- Danfma.MySheet Danfma.MySheet.Excel`
      → expected: Roughly 300 changed lines across 9 modified files plus 2 new files
      (Expressions/ImplicitIntersection.cs, Expressions/Lookup/ReferencePosition.cs). A validated reference
      diff of exactly this shape is saved at /tmp/reference-semantics-validated.patch (659 lines, generated
      from the working prototype at /tmp/mysheet-probe) — use it to cross-check the implementation, not to
      apply blindly (it lacks the doc/test items and the csharpier pass).

## Risks carried by this phase

- The 2-D intersection case is a KNOWN deviation, not a verified match. S4 settles it as #VALUE! and the prototype implements that (measured: `=A1:C3` in B2 and in C1 both #VALUE!), but Excel's @ operator most likely intersects BOTH axes and returns the cell at (formula row, formula column) when both fall inside. In the same-sheet case that is a circular reference anyway, so the deviation is usually invisible; the exposed case is `=Sheet2!A1:C3` in Sheet1!B2, where Excel would give Sheet2!B2 and MySheet gives #VALUE!. ImplicitIntersection.Intersect is deliberately structured so this becomes one extra branch before its final `return Error.Value` — do not restructure it into a two-case if/else.
- UnionReference is deliberately left at #VALUE! at the boundary (ImplicitIntersection's `default` arm) and in ROW/COLUMN (ReferencePosition's `_ =>` arm). That preserves today's measured behaviour for `=(A1:A3,B1:B3)` in a cell and for ROW((A1:A3,B1:B3)), so it is a no-change rather than a regression — but if Excel intersects a multi-area reference area-by-area, MySheet is wrong here and the fix would touch both files.
- ROW/COLUMN in an ARRAY position covers only `[NameReference or Reference]`. A reference-returning FUNCTION argument stays scalar: measured `SUM(ROW(INDIRECT("A1:A3")))` = 1 where Excel gives 6, and the same for ROW(OFFSET(..))/ROW(INDEX(..))/ROW(CHOOSE(..)). Extending the arm to Function would make ArrayEvaluation.Probe resolve — i.e. EVALUATE — that function's arguments, then TryBuildOperand resolve them a second time; for a volatile inside INDEX's row argument that is two different draws, one discarded. Do not widen the pattern without first adding a per-evaluation resolution memo.
- ReferencePosition's failure path evaluates the argument a SECOND time to recover its error (that is how ROW(NoSuchName) becomes #NAME? and ROW(INDIRECT("zz")) becomes #REF!). For a volatile argument that is a second draw whose value is discarded — harmless for an error result, but it means ROW/COLUMN of a failing volatile is not a single evaluation. It also makes ROW inconsistent with ROWS/COLUMNS/AREAS, which return 1 for an unresolvable argument (Rows.cs:29 `: 1.0`) instead of propagating. If that inconsistency is unacceptable, drop the error-propagation half of the item and accept #VALUE! everywhere.
- The boundary now runs NamedReferences.CaptureValue (a 4-case type switch) plus one byte comparison on EVERY cell evaluation at Workbook.cs:346. It is on the MISS path only — a cache hit returns from SheetValueStore before EvaluateCell is reached (Workbook.cs:183-186, :264-267) — so the cost is per first evaluation, not per read. Unmeasured, though: if a regression shows up, benchmarks/Danfma.MySheet.Benchmark/SheetBenchmarks.cs is the place to confirm it.
- Fix B makes a bare range a POSITION-DEPENDENT formula: `=A1:A4` measured 10 in C1 and 30 in C3. Any future code that assumes two cells with identical expressions have identical values — a formula-text-keyed value cache, an expression-level memo, a shared-formula group sharing one computed result — becomes wrong. Nothing does that today (IsFullyAnchored at AnchoredFormulaSupport.cs:26 is consulted only by WorksheetStreamLoader.cs:267 to pick a PARSE strategy, and each slave still evaluates through its own EvaluateCell/CellId), which is why the comment addendum at :37 is an item.
- Fix C's SameShape check changes an EXISTING result: `SUMPRODUCT(A1:A3,A1:C1)` measured 25 before and #VALUE! after. It is the Excel-documented answer (MathAggregateTests.cs:83-84 already quotes the dimension rule) and no existing test covered it, but it is a behaviour change riding inside a `fix` commit — if the release notes must be precise, split it into its own `fix(eval): SUMPRODUCT rejects mismatched dimensions` commit.
- The Phase-3 table node gets the ROW/COLUMN array path for free ONLY if it derives from `Reference` and its TryResolveReference yields a concrete RangeReference. If Phase 3 instead makes it a `Function`, the `[NameReference or Reference]` arms in ArrayEvaluation.Probe/TryBuildOperand miss it and `ROW(T[Col])` degenerates to a scalar — the corpus formula then silently returns a single row number instead of the row array. This is a contract this phase depends on, stated here so Phase 3 does not break it unknowingly.

## Open questions owned by this phase

- Does Excel's implicit intersection of a 2-D range use BOTH axes (returning the cell at the formula's row AND column when both fall inside) or is it single-axis only? Settle it in real Excel with `=Sheet2!A1:C3` typed into Sheet1!B2 — Sheet2!B2's value means both axes, #VALUE! means single-axis. This is the one case where the approved S4 paraphrase may not match Excel, and it is a one-branch change in ImplicitIntersection.Intersect.
- Does Excel intersect a MULTI-AREA (union) reference, e.g. `=(A1:A3,C1:C3)` typed in B2? If it returns A2 (first area intersected) or C2, the `default` arm of ImplicitIntersection and the `_ =>` arm of ReferencePosition both need a UnionReference case; if it returns #VALUE!, the current no-change is correct.
- Is the intersection genuinely sheet-independent for a cross-sheet reference? The prototype implements positional (`=Sheet1!A1:A3` in Sheet2!C2 → Sheet1!A2 = 0, measured) on the reasoning that the @ operator is defined on row/column NUMBERS, but I did not verify it in Excel and it is the one rule that, if wrong, would make every cross-sheet defined name in a cell return #VALUE! instead.
- ~~Should ROWS/COLUMNS/AREAS be brought in line with ROW's new error propagation?~~ DECIDED: yes — see "Scope addition" and item 27.
- SUMPRODUCT's dimension check is now orientation-aware for the rectangular backings but still count-only when either side is a union or a defined name taking the materialized fallback (PositionalRange.Rows == 0). Making a NameReference-resolved range report its shape means resolving it inside PositionalRange.Open — worth doing, but it changes a shared helper the *IFS family also uses, so it was left out of this phase.
- `=IF(TRUE,A1:A3,B1)` in a cell is still #VALUE! (measured) because If.Evaluate returns the taken branch's own Evaluate rather than routing through NamedReferences.CaptureValue, while Excel intersects it to A2. The same gap exists for IFS/SWITCH. Routing IF's branch result through CaptureValue is a two-line fix but changes IF's behaviour INSIDE functions as well (SUM(IF(TRUE,A1:A3)) would start expanding the range), so it needs its own decision rather than being smuggled in here.

## Phase Summary

_(write when phase completes)_
