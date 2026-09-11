# Phase 5: Structured-reference resolution and cross-cutting graph integration

Status: Not started   <!-- Not started | In progress | Complete -->

Part of [Structured table references, AGGREGATE, and the blocking reference-semantics gaps](../structured-table-references-and-aggregate.md) — **read that master plan first**: it carries the governing principle P0, the settled scope S1-S8, the repo-specific rules (TDD, test commands, gates, the union-tag coordination hazard) and the cross-phase open decisions. This file assumes them.

Dimension key: `resolution-and-graph`. Design dependencies: `lexer-parser`, `table-model-registry`. Adversarial verifier verdict: **needs-revision** (1 blocker, 5 majors, folded in below).

Line numbers in this file were accurate when written and several cited files have changed since. Anchor edits on member and constant names, and re-read before editing.

## Design decision

Decision: `TableReference` is a `Reference` whose ONE resolution primitive is `TryResolveRange(Workbook, out
RangeReference?, out Error)` — pure geometry over the table registry, needing no cell, no delta and no LET
scope — and every cross-cutting site calls that same primitive; `TryResolveReference` hands back the CONCRETE
`RangeReference` and `Evaluate` returns `ComputedValue.Reference(thatRange)`, mirroring
DynamicRange.cs:42-47/:53-56. I measured that this single shape lights up 23 consumers for free (SUM/COUNT/AVE
RAGE/MAX/COUNTA/ROWS/COLUMNS/AREAS/ISREF/INDEX/OFFSET/VLOOKUP/LOOKUP/MATCH/XLOOKUP/SUBTOTAL/SUMPRODUCT/SMALL/L
ARGE/MEDIAN/COUNTIF/SUMIF and LET/CHOOSE/unary-+) with zero edits — including `NamedReferences.CaptureValue`,
which needs NO new arm because `Evaluate` already yields a reference VALUE (the issue-#8 trap only bites nodes
that return #VALUE!). Four sites do NOT work and are the real content of this phase, each proven by a measured
wrong answer, not by reasoning: `ReferenceGuard` (a table on a removed sheet makes `SUBTOTAL(9,T[c])` THROW
KeyNotFoundException and `COUNTA` return 3 instead of #REF!), `ArrayEvaluation.Probe` (`COUNT((T[c]<>"")*1)` →
1 instead of 3; `MIN(IF(T[c]>0,T[c]))` → 0 instead of 5 — silent wrong numbers, the corpus formula's exact
shape), `DependencyExtractor` (`default: return;` at :216 loses the dependency without marking AlwaysDirty),
and `Row`/`Column` (#VALUE!). Failure semantics: unknown table → #NAME? (NameReference.cs:26), known table +
unknown column / missing header-or-totals row / empty data body → #REF! (Excel's own `[#Ref]` repair marker),
carried out of the ONE primitive as an `Error` so every consumer reports the same code. `Evaluate` must NOT
implement S4's implicit intersection — confirmed the only place with the formula cell's own row/column is the
private `Workbook.EvaluateCell` (Workbook.cs:328), which all three read paths funnel through
(GetCellValue:199, GetCellValueDense:282, GetCellValueOverflow:311) and which `CachedCellValue.TryFrom` :64-67
already prevents warm-start from bypassing.

## Blocking corrections — the design as written was WRONG here. Apply these first.

- [ ] **B1.** Item 21's test `Count_OfComparisonOverAnUnresolvableTable_IsRef` — "(ErrorOperand propagates #REF! rather than a #VALUE! shape mismatch)" — and verification #4's "failed: 0".
      *Measured evidence:* COUNT structurally cannot report an array element error.
      Danfma.MySheet/Expressions/Statistical/Count.cs:20-22: "// COUNT only tallies numeric values and, unlike
      SUM, never propagates errors." then `NumericAggregation.Fold(Arguments, context, ref fold);` — the
      `Error?` return is DISCARDED. Measured under /tmp/verify-resolution-and-graph (C1='=1/0', C2=4, C3=7):
      `COUNT((C1:C3<>"")*1)` => 2 (an error element in the mini-CSE stream is silently dropped by
      AddReferenced, NumericAggregation.cs:242-245), `COUNTA((C1:C3<>"")*1)` => 1, and only
      `SUM((C1:C3<>"")*1)` => #DIV/0!. So `COUNT((BadTable[Valor]<>"")*1)` will return 0.0, not #REF! — and
      #REF! vs the #VALUE! it is meant to distinguish are BOTH 0 through COUNT, so the test cannot pass and
      could not prove its claim even if inverted. Verification #4 (`--treenode-filter
      "/*/*/MiniCseConsumerTests/*"` expecting failed: 0) therefore fails after a fully correct
      implementation.
      *Correction:* Assert the ErrorOperand propagation through SUM, not COUNT:
      `Sum_OfComparisonOverAnUnresolvableTable_IsRef` = `=SUM((BadTable[Valor]<>"")*1)` → #REF! (SUM
      propagates Fold's first error; measured #DIV/0! for the identical shape). If a COUNT case is kept, its
      expectation is 0.0 and the item must say so.

## Major corrections

- [ ] **M1.** Summary + item 6 rationale: failure is "carried out of the ONE primitive as an `Error` so every consumer reports the same code", and item 6 "is also what makes VLOOKUP/INDEX/ROWS/COLUMNS report #NAME? for an unknown TABLE ... (VLookup.cs:20-24, Index.cs:33-37, Rows.cs:25-29)".
      *Evidence:* INDEX and OFFSET never call ReferenceGuard. `grep -rn "ReferenceGuard.MissingSheet"
      Danfma.MySheet/` returns 40+ call sites; Index.cs and Offset.cs are ABSENT. Index.cs:32-38 returns
      `ComputedValue.Error(Error.Ref)` directly when `NamedReferences.TryResolveReference` fails;
      Offset.cs:83-89 likewise `return Error.Ref;`. Measured with an unresolvable reference argument:
      `=INDEX(NoSuch,1,1)` => #REF!, `=OFFSET(NoSuch,1,0)` => #REF!, while `=SUM(NoSuch)` => #NAME? and
      `=SUBTOTAL(9,NoSuch)` => #NAME?. So after item 6, `INDEX(NoSuchTable[Col],1,1)` and
      `OFFSET(NoSuchTable[Col],1,0)` report #REF! while
      SUM/VLOOKUP/ROWS/COLUMNS/SUBTOTAL/COUNTIF/MATCH/XLOOKUP report #NAME?. The cited `Index.cs:33-37` is not
      reached via the guard at all.
      *Correction:* Either (a) add `ReferenceGuard.MissingSheet(Arguments[0], context)` to Index.Evaluate and
      Offset's base resolution, or (b) give TableReference an internal `TryResolveReference(context, out
      reference, out Error error)` overload the reference-consuming functions use, or (c) drop the uniformity
      claim, correct the Index.cs citation, and pin the ACTUAL per-consumer codes in item 19's matrix
      (INDEX/OFFSET → #REF!).
- [ ] **M2.** No item adds a FormulaWriter arm for TableReference, and risk #3 nonetheless assumes one exists ("FormulaWriter will canonicalize the composite form to the short one").
      *Evidence:* Danfma.MySheet/Parsing/FormulaWriter.cs:252-254: `default: throw new
      NotSupportedException($"No Excel formula rendering for node '{expression.GetType().Name}'.");`. That
      path is reachable from Danfma.MySheet.Excel/ExcelExport.cs:189 `formula =
      expression.ToFormula(sheet.Name);` (every formula on SaveAsExcel) and
      Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs:562 `return ComputedValue.Text("=" +
      expression.ToFormula(sheetName));` (FORMULATEXT of a cell whose formula contains a structured
      reference). Unlike item 5, which explicitly says "Require from table-model-registry (add if that phase
      has not)", the spec never states this requirement for lexer-parser. None of the ten verifications
      detects it: verification #1 (Release build) compiles fine because the default arm exists, and no test in
      the item list un-parses a table formula. Note the asymmetry this creates with item 17, which adds
      FORMULATEXT support for a table ARGUMENT while FORMULATEXT of a table FORMULA still throws.
      *Correction:* Add an explicit `Require from lexer-parser: a FormulaWriter.Write arm for TableReference
      (with the '-escape rules mirroring IsSimpleSheetName/WriteSheetQualifier)` item, plus one verification
      that round-trips `=SUM(Tabela1[Valor])` through `ToFormula` and one that exports a table-bearing
      workbook via SaveAsExcel.
- [ ] **M3.** Verification #9: "`grep -c 'MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs` expected 323 (322 existing + TableReference at tag 322)".
      *Evidence:* Measured on the clean checkout: `grep -c 'MemoryPackUnion'
      Danfma.MySheet/Expressions/Expression.cs` => **323** already, because Expression.cs:14 is a prose line
      containing the word ("// MemoryPackUnion tags are APPEND-ONLY: never renumber..."). `grep -c
      '^\[MemoryPackUnion' ...` => 322. So the check is satisfied BEFORE the change and FAILS (324) after it —
      inverted, and vacuous as a guard on the append-only contract.
      *Correction:* Use `grep -c '^\[MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs` expecting 323,
      and add `grep -n 'MemoryPackUnion(32[0-9]' ...` to show 319..322 contiguous.
- [ ] **M4.** S4 (approved, non-negotiable: "No ComputedValueKind.Reference may escape as a cell value") is identified but owned by nobody — no item, and dependsOn lists only lexer-parser and table-model-registry.
      *Evidence:* Measured today: a cell holding `=Rng` (a DefineName'd range) reads back with
      `GetCellValue("Data","D3").Kind == ComputedValueKind.Reference`. Workbook.cs:328-365 (`EvaluateCell`)
      has exactly two coercions — missing sheet → #REF! (:339-342) and `expression is not BlankValue &&
      value.Kind == Blank` → 0 — and NO Reference arm. Item 3 makes `TableReference.Evaluate` return
      `ComputedValue.Reference(...)`, so a bare `=Tabela1[Valor]` in a cell lands in the dense store as
      Reference, i.e. the exact state S4 forbids, and the phase adds no item for it.
      *Correction:* Either add an item implementing S4's implicit intersection in EvaluateCell (it is the only
      place all three read paths funnel through), or add the owning phase to `dependsOn` and say explicitly
      that until it lands a bare `=Table[Col]` reads back as Reference-kind.
- [ ] **M5.** The cross-cutting completeness list omits `Workbook.RestoreComparers` and the LAST-serialized-member invariant, even though S7 commits to "the new third Workbook member".
      *Evidence:* Workbook.cs:143-152 `[MemoryPackOnDeserialized] private void RestoreComparers()` rebuilds
      ONLY `Sheets` and `DefinedNames` with `StringComparer.OrdinalIgnoreCase` (and tolerates a null
      DefinedNames). Workbook.cs:119-122 states `DefinedNames` "MUST stay the LAST serialized member of
      Workbook so the schema is append-only". Item 9 is the phase's only Workbook.cs edit and touches neither.
      A `Tables` dictionary added as the third member without a RestoreComparers arm deserializes with the
      default case-SENSITIVE comparer, so `Tabela1[Col]` stops resolving in a loaded workbook (Excel table
      names are case-insensitive), and a pre-Tables file leaves it null → NullReferenceException in item 2's
      `workbook.Tables.TryGetValue`.
      *Correction:* Add a `Require from table-model-registry` item: extend RestoreComparers with the null-
      tolerant OrdinalIgnoreCase rebuild of Tables, and relocate/rewrite the "LAST serialized member" comment
      onto whichever member is now last.

## Implementation items

## CONTROLLER RULING (2026-09-10) — items 1-4 are DELETED and this phase depends on Phase 4's T1

Phase 4 item 3 and this phase's items 1-4 create the SAME file, with different enums and the same wrong union
tag. Resolved in `phase-4-lexer-parser.md`'s "CONTROLLER RULINGS on the three pre-dispatch decisions", which is
binding for both phases. In short:

- **This phase's TEXT for items 1-3 wins and moves into Phase 4's first task (T1)** — it is better reasoned and
  its reasoning is measured where Phase 4's item 3 only asserts. Do not re-derive it; carry it verbatim.
- **The enum keeps Phase 4's SIX members** (`Data = 0, All, Headers, Totals, HeadersAndData, DataAndTotals`),
  because the specifier pairs are legal Excel — measured, `COUNTA(Tabela1[[#Headers],[#Data]])` = 12 and
  `SUM(Tabela1[[#Data],[#Totals]])` = 270 with a totals row — and a four-member `TableItem` cannot represent
  what Phase 4's grammar produces. Keep `Data = 0` for the reason item 1 gives.
- **The union tag is 327**, not 322: 322 is `Aggregate` (Phase 2) and 323-326 are Phase 7's four producers.
  Count `Expression.cs` with `^\[MemoryPackUnion` at implementation time.
- **Item 5's `TryGetRegion` gains two arms and NO new guard.** `HeadersAndData` is
  `(Left..Right, TopRow..dataBottom)`, `DataAndTotals` is `(Left..Right, dataTop..BottomRow)`, and when the row
  a pair names is absent the region simply SHRINKS to the data body rather than erroring — measured,
  `[[#Data],[#Totals]]` on a table with no totals row is 9 cells summing 180, not `#REF!`. Only the
  SINGLETONS error. Copying the singleton's error arm into the pair by analogy is the mistake to avoid.
  Also: `TryGetRegion` lives on `Table`, not `TableDefinition` — Phase 3 shipped `Table`.
- **Because items 1-4 leave, T1 is the ONLY thing this phase needs from Phase 4**, so the two phases run
  concurrently from T1's merge. The remaining merge constraint is a RELEASE one and it names Phase 6, not this
  phase: nothing in `Danfma.MySheet.Excel` registers a table from an xlsx yet, so no release ships before
  Phase 6.
- **Bare `=Tabela1` resolving to the data body moves INTO this phase** from Phase 4's optional item 18, and it
  is not optional here: `=Tabela1` answers 0 today with no error where Aspose answers the data body
  (`SUM(Tabela1)` = 180, measured). Whoever takes it must answer `Parser.cs:818-826`, which argues in-tree
  that `Parser.IsCellReference` is unbounded on purpose.

- [ ] **1.** Create Danfma.MySheet/Expressions/TableReference.cs: `public enum TableItem : byte { Data = 0, All = 1, Headers = 2, Totals = 3 }` and `[MemoryPackable] public sealed partial record TableReference(string TableName, string? ColumnName, TableItem Item) : Reference`. Contract with the lexer phase, stated in the doc comment: `TableName`/`ColumnName` hold the DECODED payload (the `'`-prefix escape table `'[ '] '# ''` already applied, exactly as Tokenizer.ReadQuotedName:170-174 stores the decoded text), because IndexOfColumn compares against the raw `TableColumn/@name` from the xlsx. `Item = Data` with a non-null `ColumnName` is `T[Col]`; `Item = Data` with a null `ColumnName` is `T[#Data]`. Do NOT override `IsVolatile`.
      *Files:* `Danfma.MySheet/Expressions/TableReference.cs`
      *Why:* `Data = 0` makes the overwhelmingly common `T[Col]` form serialize the enum's default byte.
      IsVolatile is deliberately NOT overridden: Indirect.cs:16 is volatile because its target is computed
      from a runtime string no dependency tracks, whereas a table's target comes from REGISTERED state, so the
      correct invalidation mechanism is a version bump (item 9), not per-pass volatility. Overriding it would
      make DependencyExtractor.Visit's `case Function` :209-212 mark every structured-reference formula
      AlwaysDirty and throw away the whole point of the static RangeDep in item 7.
- [ ] **2.** In the same file add `internal bool TryResolveRange(Workbook workbook, out RangeReference? range, out Error error)`: `workbook.Tables.TryGetValue(TableName, out var definition)` fails → `error = Error.Name`, return false; else `definition.TryGetRegion(ColumnName, Item, out var l, out var t, out var r, out var b, out error)` fails → return false; else `range = new RangeReference(new CellAddress(l, t).ToId(), new CellAddress(r, b).ToId(), definition.SheetName)` and return true.
      *Files:* `Danfma.MySheet/Expressions/TableReference.cs`
      *Why:* ONE primitive, called by all seven cross-cutting sites (items 5-8, 10-12, 15) so the #NAME?/#REF!
      mapping and the bounds arithmetic exist exactly once. It takes a `Workbook`, not an `EvaluationContext`,
      precisely because table resolution is context-free — that is what lets DependencyExtractor.Visit (which
      only has `Workbook? wb`, DependencyExtractor.cs:68) emit a real RangeDep instead of AlwaysDirty. Unknown
      table → Error.Name follows NameReference.cs:26 (Excel resolves a table name in the same name space as a
      defined name); unknown column → Error.Ref follows Excel's own repair, which rewrites a deleted column's
      specifier to `Table1[#Ref]` and shows #REF!. `new CellAddress(col,row).ToId()` is the same construction
      DynamicRange.cs:43-44 uses.
- [ ] **3.** In the same file override `TryResolveReference(EvaluationContext context, out Reference? reference)` as `{ var ok = TryResolveRange(context.Workbook, out var r, out _); reference = r; return ok; }` and `Evaluate(EvaluationContext context)` as `TryResolveRange(context.Workbook, out var r, out var error) ? ComputedValue.Reference(r!) : ComputedValue.Error(error)`. Add a doc comment on Evaluate stating the two invariants it must never break: (a) it returns the CONCRETE resolved range, never `ComputedValue.Reference(this)`, and (b) it must never return #VALUE! the way RangeReference.Evaluate:14 does.
      *Files:* `Danfma.MySheet/Expressions/TableReference.cs`
      *Why:* Measured, both directions. (a) A probe variant returning `ComputedValue.Reference(this)` made
      `SUM(node)` return 0 instead of 14, because ComputedValue.EnumerateValues:189-191 has a catch-all `case
      Reference reference: yield return reference.Evaluate(context)` that yields the reference value back as a
      single non-numeric element which AddReferenced (NumericAggregation.cs:246-251) silently drops. Returning
      the concrete RangeReference instead hits the `case RangeReference` arm at :165 and expands correctly.
      (b) Because Evaluate returns a reference VALUE, `NamedReferences.CaptureValue`:59-69 needs NO new arm: I
      measured `LET(x,T,SUM(x))` → 14, `SUM(CHOOSE(1,T))` → 14, `SUM(+T)` → 14, and a DefineName'd table node
      → 14, all through CaptureValue's `_ => expression.Evaluate(context)` fallback at :68. DynamicRange is
      absent from that same list for exactly this reason. See item 21 for the regression pins that keep it
      that way.
- [ ] **4.** Append `[MemoryPackUnion(322, typeof(TableReference))]` immediately after `[MemoryPackUnion(321, typeof(SharedFormulaSlave))]` at Expression.cs:352, with a one-way-boundary comment in the style of the 319-321 block (:334-340), and update the stale policy line at Expression.cs:15 from "Add new tags at 319+" to 322+.
      *Files:* `Danfma.MySheet/Expressions/Expression.cs`
      *Why:* Tags 0..321 are contiguous with no gaps, so 322 is the next free tag; the block comment at
      :334-340 is the in-repo template for stating that a new tag makes a new file unreadable by an older
      build.
- [ ] **5.** Require from table-model-registry (add if that phase has not): `internal bool TryGetRegion(string? columnName, TableItem item, out int left, out int top, out int right, out int bottom, out Error error)` on TableDefinition, computing `dataTop = TopRow + HeaderRowCount`, `dataBottom = BottomRow - TotalsRowCount`, then: All → (Left..Right, TopRow..BottomRow); Data → (Left..Right, dataTop..dataBottom), and `dataBottom < dataTop` → `error = Error.Ref`, false; Headers → (Left..Right, TopRow..TopRow), and `HeaderRowCount == 0` → Error.Ref, false; Totals → (Left..Right, BottomRow..BottomRow), and `TotalsRowCount == 0` → Error.Ref, false. THEN, if `columnName is not null`, narrow: `index = IndexOfColumn(columnName)` (OrdinalIgnoreCase); `index == 0` → Error.Ref, false; else `left = right = LeftColumn + index - 1`.
      *Files:* `Danfma.MySheet/TableDefinition.cs`
      *Why:* Row band first, then column narrowing, gives all eight S1 forms from one function with no
      branching explosion: `T[Col]` = Data+column, `T[[#Data],[Col]]` is the identical node, and
      `T[[#Headers],[Col]]`/`T[[#Totals],[Col]]`/`T[[#All],[Col]]` fall out free and Excel-correct if the
      parser accepts them. Excel definitions being matched, all documented verbatim: `[#All]` = "the entire
      table, including column headers, data, and totals"; `[#Data]` = the data rows only; `[#Headers]` = the
      header row only; `[#Totals]` = the totals row only; a bare column specifier = that column's DATA cells,
      header and totals excluded. Returning FALSE on an empty region rather than an inverted rectangle is
      mandatory, not stylistic: I measured `SUM(new RangeReference("B2","B1",...))` → 5 and `ROWS(...)` → 2,
      because RangeReference.GetBounds:99-104 silently normalizes min/max — so a zero-data-row table built as
      (top=2, bottom=1) would silently READ THE HEADER ROW.
- [ ] **6.** Add to ReferenceGuard.MissingSheet(Expression, EvaluationContext), immediately after the `case DynamicRange` arm at ReferenceGuard.cs:87-94 and before `default: return null` at :96: `case TableReference table: return table.TryResolveRange(context.Workbook, out var resolvedTable, out var tableError) ? Check(context, resolvedTable!.SheetName) : tableError;`. Widen the class doc at ReferenceGuard.cs:3-10 and the method doc at :30-36 from "a reference to a sheet that does not exist" to "a structural failure of the reference itself — a missing sheet, or a structured reference whose table/column no longer resolves".
      *Files:* `Danfma.MySheet/Expressions/ReferenceGuard.cs`
      *Why:* MANDATORY for correctness, and the failure is worse than the brief predicted. Measured against a
      table node whose resolved range names a removed sheet: `SUBTOTAL(9,T[c])` THREW an unhandled
      KeyNotFoundException — Subtotal.GatherSkippingSubtotals:70 does `workbook.Sheets[range.SheetName]` with
      the throwing indexer, safe today only because Subtotal.cs:35-38 runs this guard first; `COUNTA(T[c])`
      returned 3 where the plain-range baseline returns #REF!; `COUNT(T[c])` and `COUNTIF(T[c],">0")` returned
      0 where the baseline is #REF!. Returning `tableError` (not null) on unresolvable mirrors the
      DynamicRange arm's deliberate asymmetry at :92-94, and it is also what makes VLOOKUP/INDEX/ROWS/COLUMNS
      report #NAME? for an unknown TABLE instead of the #REF!/1.0 their own `TryResolveReference`-false paths
      would give (VLookup.cs:20-24, Index.cs:33-37, Rows.cs:25-29).
- [ ] **7.** Add to DependencyExtractor.Visit, immediately after the `case UnionReference` arm (DependencyExtractor.cs:111-116) and before `case DynamicRange` at :118: `case TableReference table: if (wb is null || !table.TryResolveRange(wb, out var tableRange, out _)) { scan.AlwaysDirty = true; return; } Visit(tableRange!, scan, wb, resolving, deltaRow, deltaColumn); return;`. Extend the class doc's "Estático" paragraph (:46-48) to list TableReference and the "Always-dirty" paragraph (:50-53) to include a structured reference that does not resolve.
      *Files:* `Danfma.MySheet/DirtyGraph/DependencyExtractor.cs`
      *Why:* `default: return;` at :216-217 contributes NOTHING and does not set AlwaysDirty — the exact
      silent lost-dependency bug the class doc at :40-44 says is the one unacceptable failure. Re-dispatching
      `Visit` on the resolved RangeReference reuses the `case RangeReference` arm at :91-103 verbatim (zero
      duplicated corner arithmetic), the same re-dispatch pattern Subtotal.cs:181-186 uses for anchored nodes.
      AlwaysDirty on unresolvable mirrors ResolveName:258-262. The resolved range is absolute, so passing the
      ambient delta through is inert — correct, since a structured reference does not shift per shared-formula
      slave.
- [ ] **8.** Add `TableReference => true` to AnchoredFormulaSupport.IsFullyAnchored's switch, next to `NameReference => true` at AnchoredFormulaSupport.cs:37, with the same rationale sentence adapted (resolved by name against the workbook's table registry, independent of the slave's position).
      *Files:* `Danfma.MySheet/Parsing/AnchoredFormulaSupport.cs`
      *Why:* Without an arm the node hits `_ => false` at :62, forcing every shared-formula group containing a
      structured reference onto the legacy per-slave reparse path — safe but a silent perf cliff on exactly
      the xlsx files that have tables. Accepting it is provably correct, not optimistic: TryResolveRange (item
      2) takes only a Workbook, so its result is bit-identical for every slave in the group, which is the
      precise property `NameReference => true` at :37 is justified by. Excel parity: copying
      `=SUM(Table1[Valor])` down does not shift the reference.
- [ ] **9.** In Workbook.DefineTable (table-model-registry's method), bump the existing `_namesVersion` counter — `unchecked { _namesVersion++; }`, the same two lines DefineName does at Workbook.cs:552-555 and :585-588. Extract those three copies into `private void BumpNamesVersion()`. Widen the field comment at Workbook.cs:126-133 and the RecalculationEngine.IsStale comment at RecalculationEngine.cs:181-184 from "DefinedNames mutations" / "a defined name was (re)defined" to "workbook-level definitions (names and tables)".
      *Files:* `Danfma.MySheet/Workbook.cs`, `Danfma.MySheet/RecalculationEngine.cs`
      *Why:* Item 7 BAKES the table's resolved rectangle into the reverse dependency graph at build time —
      which is verbatim the situation RecalculationEngine.cs:181-184 says NamesVersion exists for ("a name
      redefinition touches no Sheet at all — DependencyExtractor.ResolveName bakes the name's definition into
      the graph at build time — so it needs its own version check"). Without the bump, resizing or repointing
      a table leaves a stale graph and a stale value with no eviction. Sharing the counter rather than adding
      `_tablesVersion` keeps IsStale:186-210 a single comparison and only over-invalidates, which
      DependencyExtractor.cs:40-44 states is the safe direction.
- [ ] **10.** **Phase 11a landed the `NameReference` arm exactly where this `TableReference` arm goes** — in `Probe` immediately after `case OpenRangeReference`, with its mirror in `TryBuildOperand` — so re-read both switches (the line anchors below have moved) and put the table arm beside it rather than deriving the shape again; and if `TableReference` does not derive from `Reference`, `ArrayEvaluation.IsBareReferenceNode` (`expression is Reference or NameReference`) is the SINGLE line to extend, because it is the one predicate all four top-level gates share (`TryStream`, `Index.TryResolveReference`, `NumericAggregation.Fold`'s `default:` arm and the criteria family's `PositionalRange.RejectComputedArray`) — fifteen measured top-level shapes regress if a gate is hand-written `is not Reference` instead. Add to ArrayEvaluation.Probe, next to the `case AnchoredRangeReference` arm at ArrayEvaluation.cs:135-136: `case TableReference: return (true, true);` and, next to `case Row { Arguments: [AnchoredRangeReference] }` at :146-147: `case Row { Arguments: [TableReference] }: return (true, true);`. Comment that a table's `ref` is always a bounded A1 rectangle, so the OpenRangeReference cost guard at :139-140 can never apply, and that resolution failure still yields an ARRAY (item 11) so the documented `IsArrayEligible ⇒ build succeeds as an array` invariant at :114-117 holds.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* MANDATORY for correctness and the single highest-value item in this phase. Measured with the node
      absent from Probe: `COUNT((T[c]<>"")*1)` → 1 (baseline 3) and `MIN(IF(T[c]>0,T[c]))` → 0 (baseline 5) —
      SILENT WRONG NUMBERS, not errors, because the default arm at :198-199 makes the node an opaque scalar
      broadcast (:477-479) whose `ComputedValue.Reference` compares as a single non-text value. The corpus
      formula `(ROW(T[c])-ROW(INDEX(T[c],1,1))+1)/((T[c]<>"")*(T[c]<>0))` measured #VALUE! end-to-end. Unlike
      NameReference, a TableReference can be decided context-free because it ALWAYS denotes a rectangle and
      never a scalar, so `(true, true)` cannot mis-broadcast.
- [ ] **11.** Add to ArrayEvaluation.TryBuildOperand, next to the `case AnchoredRangeReference` arm at :435-437: `case TableReference table: operand = table.TryResolveRange(context.Workbook, out var tableRange, out var tableError) ? BuildRange(tableRange!, context) : new ErrorOperand(tableError); return true;` and, next to `case Row { Arguments: [AnchoredRangeReference …] }` at :455-462, the TableReference twin building a RowNumbersOperand from `tableRange.GetBounds()` (or an ErrorOperand). Add `private sealed class ErrorOperand : ArrayOperand` after ScalarOperand (:232-244) with `IsArray => true`, `Rows => 1`, `Columns => 1`, and `At(index, rows, columns) => ComputedValue.Error(_error)` that DELIBERATELY skips the shape-mismatch check every other operand applies.
      *Files:* `Danfma.MySheet/Expressions/ArrayEvaluation.cs`
      *Why:* Resolving first and delegating to the existing BuildRange/RowNumbersOperand is strictly cleaner
      than a new operand type — it is the identical `anchoredRange.ToRangeReference(context)` normalization
      already used at :436 and :456, so no per-cell reading logic is duplicated. ErrorOperand must report
      IsArray=true or `IsArrayEligible` (item 10) would be true while the build returns not-array, breaking
      the no-double-evaluation guarantee NumericAggregation.cs:107-112 documents and MiniCseVolatileTaintTests
      pins. Skipping the mismatch check is the point: `(BadTable[c]<>"")*A1:A3` must propagate #REF! at every
      position, not be masked into #VALUE! by ArrayOperand.At's dimension rule (:226-228).
- [ ] **12.** Change the mini-CSE gate at NumericAggregation.cs:110 from `ArrayEvaluation.IsArrayEligible(argument)` to `argument is not Reference && ArrayEvaluation.IsArrayEligible(argument)`.
      *Files:* `Danfma.MySheet/Expressions/NumericAggregation.cs`
      *Why:* One line, zero behaviour change today, and it permanently closes a hijack class. Verified: this
      is the ONLY IsArrayEligible call site lacking the guard — Index.cs:23, Index.cs:171 and
      OrderStatistics.cs:507 all already write `is not Reference`, which is exactly why I measured
      `INDEX(T[c],2,1)` → 0 and `SMALL(T[c],1)` → 0 correctly on the concrete-range path. Without this line,
      item 10 would divert a bare `SUM(Table[Col])` from RangeReference.ExpandComputedValues (column-major,
      inline dense-store hit, RangeReference.cs:265-268) to RangeOperand.At (row-major, div/mod per element),
      changing both the scan order that decides which cell error wins and the hot-path cost. No existing
      Reference reaches this default arm array-eligible: RangeReference/OpenRangeReference/CellReference/Union
      Reference/AnchoredCellReference/AnchoredRangeReference all have their own arms at :56-101, and
      DynamicRange is not in Probe.
- [ ] **13.** Add a general reference fallback arm to Row.Evaluate (Danfma.MySheet/Expressions/Lookup/Row.cs:13-35) and Column.Evaluate (LookupFunctions.cs:279-297), placed AFTER the `[] when context.CellId is { } id` arm and BEFORE `_ => ComputedValue.Error(Error.Value)`: `[var single] when NamedReferences.TryResolveReference(single, context, out var resolved, boundOpenRanges: false) => resolved switch { CellReference c => Number(CellAddress.Parse(c.Id).Row), RangeReference r => Number(r.TopRow), OpenRangeReference o => Number(o.RowMin ?? 1), _ => Error(Error.Value) }` (Column: `.Column`, `LeftColumn`, `o.ColMin ?? 1`).
      *Files:* `Danfma.MySheet/Expressions/Lookup/Row.cs`, `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs`
      *Why:* This is S8, written generally rather than as a TableReference-specific arm, so it fixes every
      measured gap at once: `ROW(T[c])` → #VALUE! and `ROW(Rng)` for a defined name → #VALUE! (both measured
      this session), plus ROW(INDEX(...)) / ROW(OFFSET(...)) / ROW(INDIRECT(...)) from the brief. It cannot
      regress anything: the existing typed arms at Row.cs:15-28 still match first, and `ROW(5)`/`ROW("A1")`
      still fall to #VALUE! because TryResolveReference returns false. `boundOpenRanges: false` copies
      Rows.cs:17-21 so ROW never triggers a populated-extent sheet scan; `RowMin ?? 1` matches Excel's
      ROW(A:A) = 1. It works for `ROW(INDEX(T[c],1,1))` — the corpus anchor — because
      Index.TryResolveReference:171 gates its array path on `Arguments[0] is not Reference`, so a
      TableReference argument takes the concrete-range path there (read, not assumed).
- [ ] **14.** Add `case TableReference table:` arms to NumericAggregation.Fold (after the AnchoredRangeReference arm at :87-93) and FoldA (after :197-203) that resolve via TryResolveRange and `foreach (var value in tableRange!.ExpandComputedValues(context)) AddReferenced(value, ref fold, ref error);` (AddReferencedA in FoldA), recording `error ??= tableError` on failure.
      *Files:* `Danfma.MySheet/Expressions/NumericAggregation.cs`
      *Why:* PERF/consistency only — measured correct without it (SUM 14, COUNT 3, AVERAGE 4.667, MAX 9,
      COUNTA 3), because the default arm at :126-139 already routes a Reference-kind value through
      EnumerateValues → AddReferenced, which is the right Excel rule. The arm buys the allocation-free struct
      RangeValueSequence enumerator instead of a boxed IEnumerable, mirroring the AnchoredRangeReference arms
      added for the identical reason. Safe to defer if the phase runs long.
- [ ] **15.** In Workbook.TryGetRangeSnapshot (RangeValueCache.cs:767), before the `range is not (RangeReference or OpenRangeReference)` gate, normalize: `if (range is TableReference table && table.TryResolveRange(this, out var tableRange, out _)) { range = tableRange!; }`. Do NOT generalize the normalization to all Reference subclasses.
      *Files:* `Danfma.MySheet/RangeValueCache.cs`
      *Why:* PERF only, and it corrects the brief's assumption: the snapshot cache does NOT benefit
      automatically. Every probe site gates on the RAW argument (`argument is Reference reference` at
      CriteriaScan.cs:45-47, RangeValueCursor.cs:62-65, ArgumentFlattening.cs:180-182,
      OrderStatistics.cs:23-25, RangeAggregate.Memoize at :35) and passes that raw node, so a TableReference
      clears the outer gate and is then rejected by the inner type check — the whole second-use snapshot
      machinery is silently bypassed for every structured reference. Normalizing at this ONE choke point fixes
      all six call sites including Memoize, and keys the entry on the resolved RangeReference so `SUM(T[Col])`
      and `SUM(Data!B2:B4)` share a snapshot. Restricting it to TableReference is deliberate:
      DynamicRange.TryResolveReference EVALUATES its endpoints (DynamicRange.cs:18-19), so a blanket
      normalization here would inject endpoint evaluation into a cache-admission probe.
- [ ] **16.** Add the same `TableReference` → resolved-RangeReference up-front normalization the AnchoredRangeReference lines already do, at ArgumentFlattening.cs:98-101 (ExpandComputedValues), ArgumentFlattening.cs:50-61 (FlattenComputedValues' arm), RangeValueCursor.cs:93-96 (Open) and the matching point in PositionalRange.Open (CriteriaScan.cs).
      *Files:* `Danfma.MySheet/Expressions/ArgumentFlattening.cs`, `Danfma.MySheet/Expressions/RangeValueCursor.cs`, `Danfma.MySheet/Expressions/CriteriaScan.cs`
      *Why:* PERF only — all four measured correct without it (SUMPRODUCT(T,T) → 106, COUNTIF → 2, SUMIF → 14,
      XLOOKUP → 9) via each site's `default` arm and EnumerateValues. The normalization buys the exact-size
      List hint (ArgumentFlattening.cs:109-113) and the struct enumerator instead of a boxed iterator. If more
      than one node type ends up needing it, extract `internal static Expression ToConcrete(Expression,
      EvaluationContext)` rather than copy-pasting a second `if` into four files.
- [ ] **17.** Add `TableReference` arms to IsFormula (InformationFunctions.cs:113-130) and FormulaText (LookupFunctions.cs:526-540), resolving via TryResolveRange and using the resolved range's `StartId`/`SheetName`, mirroring the AnchoredRangeReference arms already in both switches.
      *Files:* `Danfma.MySheet/Expressions/Information/InformationFunctions.cs`, `Danfma.MySheet/Expressions/Lookup/LookupFunctions.cs`
      *Why:* Measured: `ISFORMULA(T[c])` and `FORMULATEXT(T[c])` return #VALUE! where the plain-range
      baselines return TRUE and "=B2*2". Excel reads the reference's top-left cell. Lowest priority in this
      phase (neither appears in the corpus), but they are the last two syntactic switches a table node
      silently misses.
- [ ] **18.** Add to tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs a bounds matrix over one fixture table (Data!A1:B5, headerRowCount=1, totalsRowCount=1, columns Item/Valor): `Table_ColumnSpecifier_ResolvesToTheColumnsDataRows` (B2:B4), `Table_DataItem_ResolvesToAllDataRows` (A2:B4), `Table_AllItem_IncludesHeaderAndTotals` (A1:B5), `Table_HeadersItem_IsTheHeaderRowOnly` (A1:B1), `Table_TotalsItem_IsTheTotalsRowOnly` (A5:B5), `Table_DataAndColumnComposite_EqualsThePlainColumnForm` (asserts the two parse to the SAME node), plus `Table_UnknownTable_IsName`, `Table_UnknownColumn_IsRef`, `Table_TotalsItem_WithNoTotalsRow_IsRef`, `Table_ColumnWithEscapedApostrophe_MatchesTheDecodedName` (a column literally named `O'Brien`, written `Tabela1[O''Brien]`).
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/TryResolveReferenceTests.cs`
      *Why:* This file already exists as the home for the TryResolveReference contract. The composite test
      pins that `T[Col]` and `T[[#Data],[Col]]` produce an identical node, which is the canonicalization
      FormulaWriter will round-trip (see risks). The escaped-apostrophe test is the coordination point with
      the lexer's escape table and with the corpus's `INDIRECT("Overflow2["&SUBSTITUTE(D$1,"'","''")&"]")`.
- [ ] **19.** Add to tests/Danfma.MySheet.Tests/Expressions/MissingSheetReferenceTests.cs, using its existing AssertRef harness: `StructuredReference_OnRemovedSheet_IsRefThroughEveryConsumer` parameterised over `SUM(T[Valor])`, `COUNT(T[Valor])`, `COUNTA(T[Valor])`, `SUBTOTAL(9,T[Valor])`, `COUNTIF(T[Valor],">0")`, `ROWS(T[Valor])`, `VLOOKUP(1,T[Valor],1)` (register the table, then `workbook.Sheets.Remove`), plus `StructuredReference_UnknownTable_IsNameThroughEveryConsumer` and `StructuredReference_UnknownColumn_IsRefThroughEveryConsumer`.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/MissingSheetReferenceTests.cs`
      *Why:* These are the tests that fail without item 6, and one of them (SUBTOTAL) fails by THROWING
      KeyNotFoundException rather than returning a wrong value — measured this session. The file's own doc
      comment (:6-11) already frames itself as exactly this matrix: "must propagate through EVERY consuming
      function — INCLUDING the error-ignoring COUNT family — instead of throwing KeyNotFoundException or being
      swallowed as an empty range."
- [ ] **20.** Add to tests/Danfma.MySheet.Tests/Expressions/CaptureValueTests.cs: `StructuredReference_BoundToALetName_StaysARange` (`=LET(r,Tabela1[Valor],SUM(r))`), `StructuredReference_ThroughChoose_StaysARange` (`=SUM(CHOOSE(1,Tabela1[Valor]))`), `StructuredReference_ThroughUnaryPlus_StaysARange` (`=SUM(+Tabela1[Valor])`) and `StructuredReference_AsADefinedNameDefinition_StaysARange`. Add a comment stating that these pass through CaptureValue's `_ => expression.Evaluate(context)` fallback at NamedReferences.cs:68 and that they are the guard against ever changing TableReference.Evaluate to the RangeReference.Evaluate:14 #VALUE! convention.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/CaptureValueTests.cs`
      *Why:* Tests, NOT a code edit to CaptureValue — this deliberately contradicts the brief's assumption, on
      measurement: all four already return 14 with the closed list untouched, because the node's Evaluate
      yields a reference VALUE. Adding an arm to the switch at NamedReferences.cs:62-68 would be dead code.
      But the correctness of the omission depends entirely on Evaluate's return shape, which is invisible from
      CaptureValue's file, so it needs pinning — this is the issue-#8 bug class caught by test rather than by
      arm.
- [ ] **21.** Add to tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs: `Count_OfStructuredReferenceComparison_IsElementWise` (`=COUNT((Tabela1[Valor]<>"")*1)` → 3, not 1), `Min_OfIfOverAStructuredReference_IsElementWise` (`=MIN(IF(Tabela1[Valor]>0,Tabela1[Valor]))` → 5, not 0), `Sum_OfRowOverAStructuredReference_IsTheRowVector` (`=SUM(ROW(Tabela1[Valor]))`), `SmallOverTheCorpusIdiom_MatchesTheRangeForm` (the full `(ROW(T[c])-ROW(INDEX(T[c],1,1))+1)/((T[c]<>"")*(T[c]<>0))` shape asserted equal to the same formula written with the literal range), and `Count_OfComparisonOverAnUnresolvableTable_IsRef` (ErrorOperand propagates #REF! rather than a #VALUE! shape mismatch).
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/MiniCseConsumerTests.cs`
      *Why:* Each of the first four is a measured silent WRONG NUMBER today (1 vs 3, 0 vs 5, #VALUE!, #VALUE!)
      — the strongest evidence in this phase and the only failures that a user would never notice. The paired
      "same formula with the literal range" oracle avoids inventing a golden value: the range form is already
      pinned by this file's existing K1 tests.
- [ ] **22.** Add to tests/Danfma.MySheet.Tests/DirtyGraph/DependencyExtractorTests.cs, next to `Indirect_IsAlwaysDirty` at :68: `StructuredReference_IsAStaticRangeDep_NotAlwaysDirty` asserting `Scan("=SUM(Tabela1[Valor])", workbook).Ranges` contains the exact resolved `new RangeDep("Data", 2, 2, 2, 4)` with `AlwaysDirty` false, and `StructuredReference_WithUnknownTable_IsAlwaysDirty` plus `StructuredReference_WithNoWorkbook_IsAlwaysDirty` (pass `workbook: null`).
      *Files:* `tests/Danfma.MySheet.Tests/DirtyGraph/DependencyExtractorTests.cs`
      *Why:* The existing helper `Scan(formula, workbook)` already threads a workbook for the NameReference
      tests and parses in a detached Sheet, which is fine because a TableReference carries no sheet name of
      its own. The `workbook: null` case is the one that today returns an EMPTY scan with AlwaysDirty false —
      `IsEmpty` true (DependencyScan.cs:34) — i.e. a formula the dirty engine believes depends on nothing.
- [ ] **23.** Add `StructuredReference_ThroughIndirect_Resolves` and `StructuredReference_ThroughIndirect_WithARuntimeColumnName_Resolves` to tests/Danfma.MySheet.Tests/Expressions/IndirectTests.cs — the second being the corpus shape `=SUM(INDIRECT("Tabela1["&SUBSTITUTE(D1,"'","''")&"]"))` with D1 holding a column name containing an apostrophe. Both formulas must live in a real CELL and be read via `GetCellValue`, not evaluated through `Expression.Evaluate(workbook)`.
      *Files:* `tests/Danfma.MySheet.Tests/Expressions/IndirectTests.cs`
      *Why:* Confirmed from Indirect.cs:32-68 and measured: INDIRECT works for FREE —
      `ParseFormulaBody(refText, sheet)` at :58 produces the TableReference and
      `parsed.TryResolveReference(...)` at :68 resolves it, exactly as a defined name does. I measured 14 for
      `SUM(INDIRECT("Tbl"))` and for the concat form `SUM(INDIRECT("Data!"&"B2:B4"))`. The cell requirement is
      not incidental: Indirect.TryResolveReference:50-56 returns false when `context.SheetName` is null, so a
      probe using `Expression.Evaluate(workbook)` returns #REF! — I hit exactly that false negative and it
      would look like a broken feature. The apostrophe case is the round trip of the lexer's escape table
      (item 1) against SUBSTITUTE's output.
- [ ] **24.** Add `StructuredReferenceGraph_Rebuilds_AfterDefineTable` to tests/Danfma.MySheet.Tests/DirtyGraph (a new file or the existing recalculation-engine test file): build a RecalculationEngine over a workbook with `=SUM(Tabela1[Valor])`, redefine the table with a larger `ref` via DefineTable, then assert the engine reports a rebuild / the formula's value reflects the new rows.
      *Files:* `tests/Danfma.MySheet.Tests/DirtyGraph/RecalculationEngineTests.cs`
      *Why:* Item 9's bump is invisible without this test, and its absence is a stale-value bug rather than an
      error — the same failure mode RecalculationEngine.cs:181-184 documents for defined names. Also add
      `DefineName("x", "Tabela1[Valor]")` to a Workbook test: it must NOT throw, because
      HasUnqualifiedReference (Workbook.cs:593-606) reaches `_ => false` for a TableReference — correct by
      construction, since the node carries no SheetName for the sentinel `Sheet { Name = "" }` at :572 to
      stamp with the empty string.


> **Landed differently (2026-09-09, Phase 1 Task 2):** the ROW/COLUMN axis mirror this item describes was collapsed at implementation time. The live symbols are `PositionNumbersOperand(origin, PositionAxis, rows, columns)`, `ProbePosition`/`TryBuildPositionOperand`, `ResolvePositionRange` and `PositionArgumentShape` in `Danfma.MySheet/Expressions/ArrayEvaluation.cs`. `RowNumbersOperand`, `ColumnNumbersOperand`, `ResolveRowRange` and `RowArgumentShape` no longer exist — read the file, not this text.

## Verification Plan

- [ ] `dotnet build /Volumes/Work/Develop/MySheet/Danfma.MySheet.slnx -c Release`
      → expected: Build succeeded, 0 Error(s). A compile error naming TableReference in
      ReferenceGuard/DependencyExtractor/ArrayEvaluation means an item was skipped.
- [ ] `dotnet run --project /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj`
      → expected: "Test run summary: Passed!" with failed: 0. The suite is ~2000+ tests; any pre-existing
      failure must be reproduced on a clean checkout before being attributed elsewhere.
- [ ] `dotnet run --project /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/MissingSheetReferenceTests/*"`
      → expected: failed: 0. Before item 6 this filter fails, and the SUBTOTAL case fails with an unhandled
      KeyNotFoundException ("The given key 'Data' was not present in the dictionary") rather than an assertion
      — that exception is the signature of the missing ReferenceGuard arm.
- [ ] `dotnet run --project /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/MiniCseConsumerTests/*"`
      → expected: failed: 0. Before item 10/11 the new cases fail with concrete wrong numbers —
      Count_OfStructuredReferenceComparison returns 1 (expected 3) and Min_OfIfOverAStructuredReference
      returns 0 (expected 5); those two exact numbers are the measured pre-fix values, so seeing anything else
      means the fixture differs from the spec.
- [ ] `dotnet run --project /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/DependencyExtractorTests/*"`
      → expected: failed: 0, total >= 17 (14 today + the 3 from item 23). Before item 7,
      StructuredReference_IsAStaticRangeDep fails on an EMPTY Ranges collection with AlwaysDirty false.
- [ ] `dotnet run --project /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -- --treenode-filter "/*/*/TryResolveReferenceTests/*|/*/*/CaptureValueTests/*|/*/*/IndirectTests/*|/*/*/ArrayEvaluationTests/*|/*/*/MiniCseVolatileTaintTests/*"`
      → expected: failed: 0. ArrayEvaluationTests and MiniCseVolatileTaintTests are the regression guard for
      item 11: they pin the IsArrayEligible-implies-build-succeeds invariant and the evaluate-scalar-once rule
      that ErrorOperand must not break.
- [ ] `dotnet run --project /Volumes/Work/Develop/MySheet/tests/Danfma.MySheet.Excel.Tests/Danfma.MySheet.Excel.Tests.csproj`
      → expected: "Test run summary: Passed!" with failed: 0. TableInteropTests will need the rewrite the
      excel-loader phase owns; if this suite fails only inside TableInteropTests, that phase is the owner, not
      this one.
- [ ] `cd /Volumes/Work/Develop/MySheet && dotnet csharpier check .`
      → expected: No output and exit code 0 ("Formatted 0 files" style summary with no listed file). Any
      listed file must be fixed with `dotnet csharpier format .` — this is a pre-commit gate.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -c 'MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs`
      → expected: 323 (322 existing + TableReference at tag 322). A different count means a tag was renumbered
      or duplicated — the append-only contract at Expression.cs:14-16 is broken.
- [ ] `cd /Volumes/Work/Develop/MySheet && grep -n 'IsArrayEligible' Danfma.MySheet/Expressions/NumericAggregation.cs Danfma.MySheet/Expressions/Lookup/Index.cs Danfma.MySheet/Expressions/Statistical/OrderStatistics.cs`
      → expected: Every one of the four call sites shows `is not Reference` on the same or the preceding line
      (item 12). NumericAggregation.cs is the only one missing it today.

## Risks carried by this phase

- ROW/COLUMN keep a pre-existing Excel deviation that the table node inherits: Row.cs:16 returns `range.TopRow`, so `=ROW(Tabela1[Valor])` in a cell returns the table's first data row, whereas Excel computes the implicit intersection of the ROW ARRAY and returns the formula cell's own row (or #VALUE!). This is unchanged behaviour (`ROW(A1:A3)` already measures 1, not the S4 answer), it is NOT fixable at the S4 cell boundary because ROW returns a NUMBER not a reference, and item 13 does not make it worse — but it must be documented, or a corpus formula that uses bare `ROW(T[c])` outside an aggregate will be quietly off by (dataTop - formulaRow).
- ~~The same silent-wrong-array bug item 10 fixes for tables ALREADY EXISTS for defined names and is left unfixed~~ — **CLOSED 2026-09-10 by [Phase 11a](phase-11a-unblocking-slice.md)** (Rule A, `feat/unblocking-slice`): `COUNT((Rng<>"")*1)` = 3 and `MIN(IF(Rng>0,Rng))` = 5 now, each pinned EQUAL to its literal twin in `tests/Danfma.MySheet.Tests/Expressions/DefinedNameArrayEligibilityTests.cs`, so tables and defined names will NOT disagree at this phase's release. The context-free objection is resolved the way this risk predicted: `Probe`/`TryBuildOperand` already take an `EvaluationContext`, and the name arm resolves through `NamedReferences.TryResolveReference(…, boundOpenRanges: false)` to a four-way shape (Range → array; open range → refused by the cost guard; single cell or union → scalar; unresolved → opaque), so `(A1:A3>Threshold)` with a scalar-bound `Threshold` still broadcasts, measured 2.
- `T[Column]` and `T[[#Data],[Column]]` deliberately produce the IDENTICAL node (item 5's geometry), so FormulaWriter will canonicalize the composite form to the short one. FormulaWriterTests.cs holds a flat parse-then-write IDENTITY list ("SUBTOTAL(9,A1:A3)" :268), so `Tabela1[[#Data],[Valor]]` must NOT be added there — it needs a separate normalization assertion. If the lexer phase instead adds a discriminator field to keep the round trip byte-exact, every switch in this phase still works, but item 18's `Table_DataAndColumnComposite_EqualsThePlainColumnForm` must be inverted.
- Keeping the name `ReferenceGuard.MissingSheet` while item 6 makes it also return #NAME? for an unknown table leaves the method name lying about its contract. I chose the doc-comment widening over renaming to `ReferenceGuard.Structural` because the rename touches ~20 internal call sites for zero behaviour change; the naming debt is real and should be a follow-up commit, not silently accepted.
- Item 15's normalization inside `Workbook.TryGetRangeSnapshot` sits in front of the STATEFUL second-use admission (RangeValueCache.cs:757-775). Resolving the table changes the dictionary KEY from the TableReference to the RangeReference, which is the point (a table and its literal range share a snapshot) — but it also means a formula that reads `T[Col]` once and `Data!B2:B4` once now ADMITS on what each site thinks is a first read. That is the intended, benign direction, and it is the same behaviour two literal mentions of the range already have; but the RangeValueCacheEquivalenceTests harness (which diffs cached vs uncached results) is the test that would catch it if that assumption is wrong.
- `ArrayEvaluation.ErrorOperand` reports a 1x1 shape, so a doubly-broken formula like `(BadTable[c]<>"")*(A1:A3<>"")` produces #REF! per element only because At() deliberately skips the mismatch check. If a future edit "fixes" that skip to match the other operands, the error silently becomes #VALUE! — the comment in item 11 must say so explicitly or the next reader will normalize it away.
- Item 12's `argument is not Reference` guard means a bare `SUM(Tabela1[Valor])` never enters the mini-CSE, so the ErrorOperand path is only reachable from nested positions. Should that guard be reverted or forgotten, `SUM(Table[Col])` silently switches from column-major RangeValueSequence order to row-major RangeOperand order, changing WHICH cell error wins for a multi-column `[#All]` region containing two different errors. Nothing currently tests that ordering for a table.

## Open questions owned by this phase

- What does Excel actually return for `=SUM(Table1[#Totals])` when the table has NO totals row? MS documents the specifier as returning "null" when the row is absent, which I have taken to mean #REF! (consistent with ReferenceGuard's whole reason for existing — an absent region must not degrade to an empty range and a silent 0 — and with the DynamicRange precedent at ReferenceGuard.cs:92-94). But Excel may instead treat it as an empty reference and return 0. Same question for `[#Headers]` on a table with headerRowCount=0. Settle it empirically without Excel: author a fixture .xlsx IN EXCEL containing `=SUM(Table1[#Totals])` and `=Table1[#Totals]` over a table with no totals row and read the CACHED <v>/<t="e"> Excel stored next to each formula — the loader already reads that cache (WorksheetStreamLoader.cs:500-515), so the fixture is self-verifying. Neither machine in this session has Excel or LibreOffice (checked), so I could not measure it.
- Is #NAME? or #REF! correct for a table name that never existed (`=SUM(NoSuchTable[Col])`)? I chose #NAME? by analogy with NameReference.cs:26 and Excel's shared name space for tables; the opposing evidence is that Excel's OWN repair path for a broken structured reference writes the `[#Ref]` marker and yields #REF!, though that marker covers a DELETED column, not an unknown table. Both are errors, so IFERROR/ISERROR-wrapped corpus formulas are unaffected — the only observable difference is ERROR.TYPE and IFNA. Settle with the same cached-<v> fixture experiment.
- ~~Should the NameReference array-eligibility gap (see risks #2) be fixed in this feature or deferred?~~ **ANSWERED and DELIVERED 2026-09-10 by [Phase 11a](phase-11a-unblocking-slice.md)** — neither in this feature nor deferred: it shipped BEFORE this phase, so nothing here has to scope it. Three shapes are deliberately left over there and are none of this phase's business: an open-range name (the mini-CSE cost guard), a union name (the union ruling owns it) and a `LET` node in a consumer's own argument slot (Phase 7's `LET` routing).
- Can an xlsx table legitimately have ZERO data rows (`ref="A1:B1"` with headerRowCount=1)? I specified #REF! for that case, because the alternative — an inverted rectangle — silently reads the header row (measured: `ROWS(RangeReference("B2","B1"))` = 2). Excel's UI always keeps at least one blank data row, so this may be unreachable in practice, but the file format permits it and ClosedXML/third-party writers may emit it. If the excel-loader phase can guarantee it normalizes such a table at registration time, the check in item 5 becomes defence in depth rather than behaviour.
- Does the table registry need to make the BARE table name resolvable (`=SUM(Tabela1)`, which Excel treats as `Tabela1[#Data]`)? It is outside S1, and today `Tabela1` parses to a NameReference and yields #NAME? unless table-model-registry also registers the name. Leaving it as #NAME? is a visible gap against Excel; registering it means a table name and a defined name can collide, which S2 says must be prevented anyway. Flagging it so the registry phase makes the call deliberately rather than by omission.

## Phase Summary

_(write when phase completes)_
