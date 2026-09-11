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

## CONTROLLER RULINGS after the re-verification (2026-09-11) — binding

Read the re-verification appended below first; these settle its R1-R4 and its sequencing.

- **R1 — a table with ZERO data rows (header-only) is a recorded DIVERGENCE, not matched.** The oracle treats it as an EMPTY reference (`SUM` 0, `ROWS` 0, `ISREF` TRUE, `AVERAGE` `#DIV/0!`, `INDEX` `#REF!` — the full row list is in the re-verification's finding 2). MySheet has no zero-extent reference node (`RangeReference.GetBounds` normalizes an inverted pair to a real rectangle; `ArrayShaping` forbids a 0-extent operand), and Excel's own UI cannot create the shape — only ClosedXML-class writers can. So this falls under P0's "structurally cannot match" clause: `Data` on a header-only table answers `#REF!`, pinned beside the oracle's table, and recorded as sweep item 33 so the decision can be reopened. **`TryGetRegion` MUST still carry a third outcome (`Empty`) distinct from `Error.Ref`**, so that reopening it later is one arm and not a re-plumbing of the primitive. The user can override this ruling; it is stated in the ledger and in the report to the user.
- **R2 — `ReferenceGuard` gets ONE arm that checks only the RESOLVED range's sheet and returns `null` when the table does not resolve.** An unresolvable structured reference is an error VALUE and each consumer does with it what it does with any error-valued argument: `COUNT(Tabela1[#Totals])` is **0**, `COUNTA` is **1**, `SUM` is `#REF!`, `ISREF` is FALSE — measured, both modes, and identical in shape to the oracle's unknown-name column. Item 6's `tableError` short-circuit and item 19's uniform-code matrix are struck; M1's uniformity goal is inverted. The removed-sheet half of item 6 stands.
- **R3 — bare `=Tabela1` uses mechanism (a):** `ParseIdentifier` routes an identifier that is cell-shaped but NOT `Parser.IsExcelGridCellReference` to `NameReference` (the `IsCellReference` helper and `ExcelGridCellReferenceTests.IsCellReference_StaysUnbounded` are untouched); `NamedReferences.TryResolveRaw` and `NameReference.Evaluate` resolve a name found in `Workbook.Tables` to `TableReference(name, null, Data)` at evaluation time. No workbook at parse time, and `DependencyExtractor.ResolveName` plus `ReferenceGuard`'s `NameReference` arm see it for free. It lands in Phase 5's T5, AFTER Phase 4's T5 merges (shared `Parser.cs`).
- **R4 — the error-valued range slot in the criteria family goes to the SWEEP** (item 34): `COUNTIF(NoSuch,">0")` is 0 here and `#NAME?` there today, for every defined name; this phase pins the table instance with both numbers and does not fix the general rule.
- **Sequencing:** every Phase 5 task rebases onto the merged 11c and Phase 4 T1 and re-reads both `ArrayEvaluation` switches. T1 (geometry) runs alone first; T3 and T4 concurrently (no shared file); T2 beside them; T5 last.
- **Shipped already, struck from the items:** 1-4 (Ruling 1 of Phase 4), 8 (Phase 4 T4a), 9, M5 (Phase 3), 12 (Phase 11a), 13, M4 (Phase 1). M2 and M3 are Phase 4's. Fifteen items remain: 5 (re-ruled by R1 and the six-area geometry), 6 (halved by R2), 7, 10, 11 (shrunk — `SingletonArrayOperand` is the `ErrorOperand`; no `Row` twin), 14-18, 19 (rewritten), 20-24.

## Re-verification against main @ a99ab43 (2026-09-11)


Scope: `plans/structured-table-references-and-aggregate/phase-5-resolution-and-graph.md` (24 items, B1, M1-M5), read against the tree at `main` HEAD `a99ab43` with Phases 1, 2, 3, 7, 8, 9, 10, 11a and 11b merged and Phase 11c in flight (`/Volumes/Work/Develop/MySheet-11c`, `feat/array-bindings`, not touched). READ-ONLY: nothing in the repo was modified; every experiment ran under `…/scratchpad/p5-reverify/` (an Aspose probe copied from `/tmp/aspose-probe-fable`, and a console project referencing `Danfma.MySheet.csproj` by path).

Measured baselines on `main` today: Release build `--no-incremental` **0 warnings / 0 errors**; core suite **1927 total / 1 failed** — the one failure is Phase 7's deliberate red pin `DynamicArrayTests.ALetBoundProducer_InACriteriaSlot_IsAKnownLimitHandedOverByPhase11a`, which Phase 11c closes; Excel suite **93 / 0**. Per class: `DependencyExtractorTests` 18, `MissingSheetReferenceTests` 58, `TryResolveReferenceTests` 10, `CaptureValueTests` 4, `MiniCseConsumerTests` 54, `IndirectTests` 7, `ArrayEvaluationTests` 20, `MiniCseVolatileTaintTests` 13, `RecalculationEngineTests` 15, `DefinedNameArrayEligibilityTests` 30, `TableRegistryTests` 24. `^\[MemoryPackUnion` count in `Expression.cs`: **327** (tags 0-326 contiguous; the policy comment at `:15` already says "327+").

Oracle: Aspose.Cells 26.6.0, fixture `Data!Tabela1` = `A1:C4`, header row `Item / Valor / Qtd`, data rows `a,10,1 / b,20,2 / c,30,3`, no totals row unless a row says so; formulas typed on `Main` at `H20+` (outside the table's rows), every row measured PLAIN and array-entered (CSE, `SetArrayFormula`). **This engine implements the array-entered rule inside a function argument and the PLAIN (implicit-intersection) rule at the cell boundary**; a row below names the mode whenever the two differ. No `TableReference` node exists on `main` (grep over `Danfma.MySheet`, `Danfma.MySheet.Excel`, `tests`: zero hits), and there is no Phase 4 branch or worktree, so Phase 4's T1 has not started.

## The findings that do the most damage unnoticed, ranked

1. **Item 6, M1 and item 19 prescribe a structural short-circuit for an UNRESOLVABLE table, and the oracle says the failure is a plain error VALUE.** The one unresolvable structured reference Aspose accepts at entry is `[#Totals]` on a table with no totals row (an unknown table or column is rejected at formula-set time: `Invalid table reference[NoSuchTable]`, `Invalid table column: NoSuchCol`). Measured, both modes agree on every row: `SUM(Tabela1[#Totals])` `#REF!`, **`COUNT` 0, `COUNTA` 1**, `COUNTIF(…,">0")` `#REF!`, `SUMIF` `#REF!`, `COUNTBLANK` `#REF!`, `SUMPRODUCT` `#REF!`, `AVERAGE`/`MAX`/`SMALL`/`MEDIAN` `#REF!`, `VLOOKUP`/`MATCH`/`XLOOKUP` `#REF!`, `INDEX(…,1,1)` `#REF!`, `OFFSET(…,0,0)` `#REF!`, `ROW`/`COLUMN`/`ROWS`/`COLUMNS`/`AREAS` `#REF!`, **`ISREF` FALSE**, `ISERROR` TRUE, `ERROR.TYPE(SUM(…))` 4, `SUBTOTAL(9)` and `SUBTOTAL(3)` `#REF!`, `AGGREGATE(9,6,…)` `#REF!`, `SUM((…<>"")*1)` `#REF!`, **`COUNT((…<>"")*1)` 0, `COUNTA((…<>"")*1)` 1**, `SUM(…*A1:A3)` `#REF!`, `COUNT(…*A1:A3)` 0, `SUM(IF(…>0,1,0))` `#REF!`, `LET`/`CHOOSE`/`+`/`INDIRECT` `#REF!`, `=Tabela1[#Totals]` bare `#REF!`, `ISFORMULA` FALSE, `FORMULATEXT` `#REF!`, `COUNT(Tabela1[Valor],Tabela1[#Totals])` 3 in either order, `SUM(Data!B2:B4,Tabela1[#Totals])` `#REF!`. That is exactly the oracle's behaviour for an unknown NAME with the code swapped (`SUM(NoSuch)` `#NAME?`, `COUNT` 0, `COUNTA` 1, `COUNTIF` `#NAME?`, `ISREF` FALSE, `INDEX`/`OFFSET`/`VLOOKUP`/`MATCH` `#NAME?`, `COUNT((NoSuch<>"")*1)` 0, `COUNTA((NoSuch<>"")*1)` 1) and for a literal error (`COUNT(1/0)` 0, `COUNTA(1/0)` 1, `COUNTIF(1/0,…)` `#DIV/0!`): **the node evaluates to an error value and each consumer does with it what it does with any error-valued argument.** Item 6's `case TableReference … : tableError` inside `ReferenceGuard.MissingSheet` would make `COUNT(Tabela1[#Totals])` `#REF!` (via `Count.cs:15-18`) and `COUNTA` `#REF!` (via `CountA.cs:14-17`) — two silent divergences in the very family the guard exists for — and item 19's `StructuredReference_UnknownColumn_IsRefThroughEveryConsumer` parameterised over `COUNT`/`COUNTA` would pin them. The arm must check only the RESOLVED range's sheet (`TryResolveRange` true → `Check(context, range.SheetName)`) and return **`null`** when the table does not resolve; the tree then already answers the oracle's rows for free: `NumericAggregation.Fold`'s `default:` (`:103-155`) evaluates the node, `AddDirect` (`:296-304`) puts the error on the channel COUNT discards; `ArgumentFlattening.FlattenComputedValues`' `default:` (`:79-105`) yields the error value COUNTA counts; `AggregateCodes.Gather`'s `default:` (`:197-215`) propagates a whole-argument error unconditionally; `ReferencePosition.Unresolved` (`:151-160`) reports the argument's own error for `ROW`/`ROWS`/`COLUMNS`/`AREAS`; `IsRef` (`InformationFunctions.cs:97-104`) answers FALSE; `Index.cs:28-34`, `Offset.cs:84-88`, `VLookup.cs:20-24` answer `#REF!`. The removed-sheet half of item 6 stands (see finding 3). M1's "every consumer reports the same code" goal is therefore inverted, not refined: `COUNT` must NOT report it at all.

2. **Open question 4 is answered, and it inverts item 5: a table with ZERO data rows is an EMPTY reference on the oracle, not `#REF!`.** Aspose accepts `ListObjects.Add(0,0,0,2,true)` (header-only, `ref="A1:C1"`), and over it (both modes unless noted): `SUM(Tabela1[Valor])` **0**, `COUNT` 0, `COUNTA` 0, **`ROWS` 0**, `COLUMNS` 1, `AREAS` 1, **`ISREF` TRUE**, `ERROR.TYPE(SUM(…))` `#N/A` (no error), `AVERAGE` `#DIV/0!`, `MIN`/`MAX` 0, `SMALL(…,1)` `#NUM!`, `MEDIAN` `#NUM!`, `INDEX(…,1,1)` `#REF!`, `MATCH`/`VLOOKUP` `#N/A`, `COUNTIF`/`SUMIF` 0, `COUNTBLANK` 0, `SUBTOTAL(9)`/`(3)` 0, `SUMPRODUCT(T[Valor])` 0, `SUMPRODUCT(T[Valor],T[Qtd])` `#VALUE!` plain / 0 CSE, `SUM(T[Valor]*2)` `#VALUE!` plain / 0 CSE, `ROWS(T[Valor]*2)` 0, `SUM((T[Valor]<>"")*1)` `#VALUE!` plain / 0 CSE, `COUNT((T[Valor]<>"")*1)` 0, `FILTER(T[Valor],…)` `#CALC!`, `SUM(OFFSET(T[Valor],0,0))` 0 and `ROWS` of it 0, `ROWS(T[Valor]:T[Qtd])` 0, `SUM(Tabela1)`/`SUM(T[#Data])` 0, `ROWS(T[#Data])` 0, `ROWS(T[[#Headers],[#Data]])` 1, `COUNTA(T[[#Headers],[#Data]])` 3, `ROWS(T[[#Data],[#Totals]])` 0, `SUM(T[#Totals])` `#REF!` (the singleton still errors), `LET`/`IF(TRUE,…)`/`INDIRECT` 0, `=Tabela1[Valor]` bare `#VALUE!` plain / 0 CSE, `ISERROR(T[Valor])` **TRUE plain / FALSE CSE**, `ROW(T[Valor])` **2 plain / 1 CSE**, `COLUMN` 2. Item 5's `dataBottom < dataTop → Error.Ref` and its rationale ("an absent region must not degrade to an empty range and a silent 0") are the opposite of the measurement; `Table.TryGetColumnRange` (`Table.cs:114-133`) already returns `false` for `DataRowCount == 0` with the comment "the caller owns the `#REF!` decision" (`:111-112`), so the decision is still open in the tree. MySheet has no zero-row reference node (a `RangeReference` always has at least one cell; `RangeReference.GetBounds:89-105` normalizes an inverted pair to a real rectangle — re-measured: `ROWS(new RangeReference("B2","B1"))` = 2, `TopRow` = 1, so item 5's "an inverted rectangle silently reads the header row" still holds). This is a controller ruling, not an implementer's call: match the oracle with an empty-reference representation (new work, touching every consumer that pattern-matches `RangeReference`), or record a deliberate divergence under P0's "structurally cannot match" clause with the table above pinned beside it. Excel's UI cannot create the shape; ClosedXML and other writers can; Phase 6's loader will meet it.

3. **Six items and two majors are already shipped by merged phases, and item 8 is Phase 4's.** Item 9 (`_namesVersion` bump in `DefineTable`): shipped in Phase 3 as one counter, `Workbook._definitionsVersion` (`Workbook.cs:155-172`), bumped at `:609-612`, `:641-644` and `:675-679`, read by `RecalculationEngine.HasDefinitionChange` (`:230`) with the sticky flag (`:87-94`), and pinned by `RecalculationEngineTests.RedefiningATable_ForcesAFullRecompute*` (`:336-396`, under `#if MYSHEET_TABLES`, which `Danfma.MySheet.Tests.csproj:15` defines) and `TableRegistryTests.DefinitionsVersion_AdvancesOnEachDefineTable` (`:340`). The symbol `_namesVersion`, `NamesVersion` and `IsStale` no longer exist. Item 12 (`argument is not Reference &&` before `IsArrayEligible` in `NumericAggregation.cs:110`): shipped in Phase 11a as `ArrayEvaluation.TryStream` (`:227-240`, gate `:233` = `!IsBareReferenceNode(expression) && IsArrayEligible(…)`), which `NumericAggregation.Fold`'s `default:` calls at `:126`; there is no `IsArrayEligible` call left in `NumericAggregation.cs` or `OrderStatistics.cs` (SMALL/LARGE go through `OrderSelection.cs:71`, `TryStream`), and `Index.cs:175-176` uses `IsBareReferenceNode`. Item 13 (S8, a general reference fallback in `ROW`/`COLUMN`): shipped in Phase 1 — `Row.cs:40` `[var only] => ReferencePosition.Row(only, context)`, `LookupFunctions.cs:298` the `COLUMN` mirror, `ReferencePosition.cs:31-52`; re-measured `ROW(Rng)` = 2 and `COLUMN(Rng)` = 2 for a defined name (the item's "`ROW(Rng)` → `#VALUE!`" is stale), and `ROW(Tabela1[Valor])` will resolve through `NamedReferences.TryResolveReference` → the virtual `TableReference.TryResolveReference` with no arm. M4 (S4 at the cell boundary): shipped in Phase 1 — `Workbook.EvaluateCell` (`:365-420`) captures through `NamedReferences.CaptureValue` (`:392`) and intersects a reference-kind result through `ImplicitIntersection.Apply` (`:400-402`); re-measured `=Rng` in `Main!H2` reads back **10, kind Number**, in `Data!D3` **20** — no `Reference` kind escapes, and `CachedCellValue.TryFrom`'s `Reference` arm (`:70-75`) is documented unreachable from the snapshot path. A bare `=Tabela1[Valor]` will follow the same route the moment `TableReference.Evaluate` yields the concrete range; the oracle's PLAIN column agrees (`#VALUE!` in `H20`, 10/20/30 over the data rows per Phase 4's measurement). M5 (`RestoreComparers` + the LAST-serialized-member comment): shipped in Phase 3 — `Workbook.cs:177-187` rebuilds `_tables` null-tolerant with `OrdinalIgnoreCase`, and the append-only comment now sits on `_tables` at `:129-141`. Item 8 (`TableReference => true` in `AnchoredFormulaSupport.IsFullyAnchored`) is Phase 4 item 11 (task T4a) verbatim, as Phase 4's finding 4 already recorded; the arm is at `:42`, not `:37`, and `_ => false` at `:67`, not `:62`. Items 1-4 are deleted by Ruling 1. What remains of the 24 is items 5 (re-ruled), 6 (halved), 7, 10, 11 (shrunk), 14-18, 19 (rewritten), 20-24 (numbers moved) — fifteen items, five of them PERF-only.

4. **Risk #1 is false on the oracle: `ROW(Tabela1[Valor])` is the table's first data row in BOTH modes.** Measured 2 / 2 on `Main` (formula at row 20, far from the table) and 2 / 2 on `Data`, and `ROW(Tabela1)` 2 / 2, `ROW(Tabela1[#Headers])` 1 / 1, `ROW(Tabela1[#Totals])` 5 / 5 with a totals row. The risk's claim that Excel "computes the implicit intersection of the ROW ARRAY and returns the formula cell's own row (or `#VALUE!`)" is not what Aspose does; `Row.cs:16`'s `range.TopRow` is the oracle's answer. Retire the bullet. (The only split is the header-only table above: 2 plain / 1 CSE — record it as an oracle mode split, not a rule.)

5. **Item 11's `ErrorOperand` already exists under another name, and its `Row` twin is unnecessary.** `SingletonArrayOperand` (`ArrayShaping.cs:78-89`) is a 1x1 array operand whose `At` returns its value at any index — byte for byte the proposed `ErrorOperand`, and since Phase 10 a 1x1 broadcasts by rule (`Broadcasting.TryProject`, `Broadcasting.cs:54-94`; `ArrayOperand.At` doc `ArrayOperands.cs:89-99`), so "DELIBERATELY skips the shape-mismatch check every other operand applies" describes the documented invariant rather than a deliberate hole, and the risk bullet about a future edit "fixing" the skip is retired. The build arm is `operand = table.TryResolveRange(context.Workbook, out var range, out var error) ? BuildRange(range!, context) : new SingletonArrayOperand(ComputedValue.Error(error)); return true;` beside `case NameReference` (`ArrayEvaluation.cs:494-514`), and the probe arm `case TableReference: return (true, true);` beside `:293-299` — ABOVE `case IArrayProducer` (`:369-370`), whose own comment at `:367-368` reserves the spot. `Row { Arguments: [TableReference] }` needs no arm: the pattern at `:306`/`:520` is `[NameReference or Reference]`, it admits a `TableReference`, and `ResolvePositionRange` (`:687-735`) resolves it through `NamedReferences.TryResolveReference` → the virtual, degrades an unresolvable one to `Scalar` so `ROW`'s own error broadcasts (oracle: `ROW(Tabela1[#Totals])` `#REF!`), and re-checks the resolved sheet at `:716-719`. Item 11's "RowNumbersOperand from `tableRange.GetBounds()`" and the orphaned "Landed differently" note at the end of the items (`PositionNumbersOperand`/`TryBuildPositionOperand`) both go.

6. **The union-tag verification is Phase 4's, and its numbers are wrong twice over.** Verification #9 says `grep -c 'MemoryPackUnion'` expecting 323; M3 corrects the pattern to `^\[MemoryPackUnion` expecting 323. Today the anchored count is **327** (322 `Aggregate`, 323-326 `Filter`/`Sort`/`Unique`/`Sequence` at `Expression.cs:353-359`), Ruling 1 makes T1 write **327**, and the check after T1 is **328**. The block-comment template for a one-way boundary is now the Phase 7 comment at `:354-355` as much as the 319-321 one at `:344-349`. Keep the check in Phase 5's plan only as "confirm T1 landed 327 and the count is 328"; the edit belongs to T1.

7. **B1 holds and is now oracle-confirmed, but its COUNTA number moved.** Re-measured on `main`: `COUNT((C1:C3<>"")*1)` with `C1` = `1/0` is **2**, `SUM` is `#DIV/0!` — B1's mechanism is intact (`Count.cs:20-22` discards `Fold`'s error channel; `AddReferenced` is `:249-266`, `AddDirect` `:296-340`). `COUNTA((C1:C3<>"")*1)` is now **3**, not the 1 B1 recorded, because Phase 7 gave `FlattenComputedValues` a `TryStream` arm (`ArgumentFlattening.cs:82-90`) and an error element is non-blank. The oracle confirms both halves of the correction on the table shape: `SUM((Tabela1[#Totals]<>"")*1)` = `#REF!` and `COUNT((Tabela1[#Totals]<>"")*1)` = 0 in both modes, so `Sum_OfComparisonOverAnUnresolvableTable_IsRef` is the right pin and a kept COUNT case expects 0.

8. **Phase 5 has ZERO docs items and four shipped sentences go false the day it merges.** `docs/workbook-and-expressions.md:993-1000` ("The structured-reference syntax is not implemented yet … So nothing in the evaluator reads a table yet: you register one to keep the model through a round trip") with pt-BR `:1045-1052`; `docs/serialization.md:344-347` ("nothing in the formula language reads a table yet … the next free tag is **327**") with pt-BR `:374-376` — the first half is Phase 4 T6a's, the "reads a table" half and the tag sentence are this phase's / T1's; `docs/excel-interop.md:66`, `:125`, `:243-246` and pt-BR `:68`, `:133`, `:264-270` ("cannot be parsed and degrades to the cached value" — T6a's, but the row at `:66` must gain what a registered table's structured reference now resolves to); and `docs/function-reference.md` rows `ROW` (`:256`), `ROWS` (`:257`), `INDIRECT` (`:252`), `ISREF` (`:282`), which enumerate the reference-producing forms each accepts and must gain the structured reference. The pt-BR `workbook-and-expressions.md` is already asymmetric around the Tables section (Phase 4 finding 7), so the twin edit differs and needs a paragraph-by-paragraph diff.

9. **Bare `=Tabela1` (Ruling 3's half) cannot be done in the parser as the ruling words it, because the parser has no workbook.** `Parser` is constructed over a `Sheet` (`Parser.cs:10`), `ExpressionParser.Parse(string, Sheet)` (`:14`) and `ParseFormulaBody(string, Sheet)` (`:39`) take no `Workbook`, and `Sheet` carries no back-reference (grep `Workbook` in `Sheet.cs`: doc comments only). `ParseIdentifier` (`:308-331`) classifies `Tabela1` as a cell at `:326` through the unbounded `IsCellReference` (`:788-812`), and `Parser.cs:818-826` argues in-tree that it is unbounded on purpose ("MySheet's grid has no ceiling and `ParseIdentifier` depends on that"), with `ExcelGridCellReferenceTests.IsCellReference_StaysUnbounded` (`:51-56`) pinning the helper. Re-measured on `main`: `=Tabela1` **0**, `=SUM(Tabela1)` **0**, `=ROWS(Tabela1)` **1**, `=Tabela1+1` **1** — the silent wrong number is real. Oracle (both modes): `SUM(Tabela1)` **66**, `ROWS` 3, `COLUMNS` 3, `COUNTA` 9, `ISREF` TRUE, `AREAS` 1, `INDEX(Tabela1,2,2)` 20, `COUNTIF(Tabela1,">15")` 2, `LET(t,Tabela1,SUM(t))` 66, `SUM(INDIRECT("Tabela1"))` 66, `SUM(Tabela1[])` 66 and stored back as `=SUM(Tabela1)`, `SUM(Tabela1*1)` `#VALUE!` in both modes (a text column in the body), `COUNT((Tabela1<>"")*1)` 0 plain / 9 CSE, `=Tabela1` bare `#VALUE!` plain (2-D) / "a" CSE (top-left). Two mechanisms exist, and the choice is the controller's: (a) tighten the classification in `ParseIdentifier` only (leave the `IsCellReference` helper and its pin alone; route an identifier that is a cell shape but NOT `IsExcelGridCellReference` to `NameReference`), then resolve a `NameReference` whose name is in `Workbook.Tables` to `TableReference(name, null, Data)` at evaluation time in `NamedReferences.TryResolveRaw` (`:116-160`) and `NameReference.Evaluate` (`:14-27`) — no workbook needed at parse time, `DependencyExtractor.ResolveName` (`:249-285`) and `ReferenceGuard`'s `NameReference` arm (`:70-82`) then see it for free, and every cell beyond `XFD` keeps parsing as a cell only if its letter run is ≤ 3, which is exactly the set `Table.ValidateName` already reserves (`Table.cs:139-148`); or (b) thread a registry lookup into the parser, which changes three public entry points. Either way the work lands in `Parser.cs`, which Phase 4's T5 owns — sequence it after T5 or fold it into T5/T6b.

10. **Phase 11c edits the two files items 10/11 edit, and merges first.** 11c items 6, 7, 9, 10 and 11 touch `ArrayEvaluation.cs` (the `NameReference` arms at `:293-299`/`:494-514`, a new `IsBareReferenceNode(Expression, EvaluationContext)` overload beside `:252-253`, a `ResolveNameShape` outcome at `:775-799`, `Choose` and `UnaryOperation{Plus}` arms) and `NameReference.cs`, `NamedReferences.cs`, `CriteriaScan.cs`, `Index.cs`. The master plan's phase table orders 11c "before Phases 4, 5 and 6". Phase 5's arms are additive and go immediately after `case NameReference` in both switches; the overload changes nothing for a `TableReference` (`expression is Reference` is true regardless of scope); the only hazard is textual, on rebase. Phase 5's brief must say "rebase onto the merged 11c and re-read both switches", as Phase 7's did for 11a.

## 1. Anchors, exhaustively

Legend: EXACT = line(s) hold what is claimed; MOVED = same content, new line; FALSE = the line does not carry the claim (see §2); GONE = symbol no longer exists.

| cited | status | current |
| --- | --- | --- |
| `DynamicRange.cs:42-47/:53-56` (design, item 3) | MOVED | `TryResolveReference` `:13-48` (endpoint resolution `:18-19`, `new CellAddress(…).ToId()` `:43-44`); `Evaluate` `:53-56` EXACT |
| `NameReference.cs:26` (`#NAME?`) | EXACT | `:26` |
| `Workbook.cs:328` `EvaluateCell`; `:328-365`; `:339-342` missing sheet | MOVED | `EvaluateCell` `:365-420`; missing sheet `:376-380`; capture `:392`; intersection `:400-402`; blank→0 `:410-413` |
| `GetCellValue:199`, `GetCellValueDense:282`, `GetCellValueOverflow:311` | MOVED | `:206`, `:296`, `:331` |
| `CachedCellValue.TryFrom :64-67` | MOVED | `TryFrom` `:35`; `Reference` arm `:70-75` (comment says the boundary intersects every reference away) |
| `DependencyExtractor.cs :216` (`default: return;`) | EXACT | `:216-217` |
| `Count.cs:20-22` | EXACT | `:20-22` (guard `:13-18` above it) |
| `NumericAggregation.cs:242-245` / `:246-251` (`AddReferenced`) | MOVED | `:249-266` |
| `Index.cs:32-38` → `Error.Ref`; `Index.cs:33-37` | MOVED | `:28-34` (`TryResolveReference` fails → `#REF!` `:33`); `TryStream` gate `:22` |
| `Offset.cs:83-89` | EXACT | `:84-88` |
| `VLookup.cs:20-24` | EXACT | `:20-24`; guard call `:13` |
| `Rows.cs:25-29` ("1.0 their own TryResolveReference-false paths would give") | FALSE | `Rows.cs:35-46`: an argument that fails to resolve reports its OWN error (`ReferencePosition.TryResolve`/`Unresolved` `:100-160`); re-measured `ROWS(NoSuch)` = `#NAME?` |
| `FormulaWriter.cs:252-254` default throw | MOVED | `:253-255` |
| `ExcelExport.cs:189` `ToFormula` | EXACT | `:189` (and `:124` for defined names) |
| `LookupFunctions.cs:562` FORMULATEXT `ToFormula` | MOVED | `:601`; the switch `:567-583`, `#VALUE!` `:587` |
| `Expression.cs:14` prose line; `:15` "319+"; `:352`; `:334-340` | MOVED/FALSE | `:14-15` says **"327+"** already; 321 is at `:352` but 322-326 follow (`:353-359`); the 319-321 comment is `:344-349`; count `^\[MemoryPackUnion` = **327** |
| `Workbook.cs:143-152` `RestoreComparers`; `:119-122` LAST member | MOVED | `:177-187` (rebuilds `_tables` null-tolerant); the LAST-member comment is on `_tables` at `:129-141` |
| `Tokenizer.ReadQuotedName:170-174` | MOVED | `ReadQuotedName` `:159`; the `''` escape `:168-173` |
| `Indirect.cs:16` `IsVolatile` | EXACT | `:16` |
| `DependencyExtractor.Visit case Function :209-212` | EXACT | `:207-214` |
| `DependencyExtractor.cs:68` (`Workbook? wb`) | EXACT | `Extract` `:68`, `Visit` `:78` |
| `ComputedValue.EnumerateValues:189-191`, `:165` | EXACT (path) | file is `Danfma.MySheet/ComputedValue.cs`, not `Expressions/`; `:156`, `:165`, `:189-190` |
| `NamedReferences.CaptureValue:59-69`, `:68`, `:62-68` | MOVED | `CaptureValue` `:59-80`, default `_ => expression.Evaluate(context)` at **`:79`**; `SharedFormulaSlave` arm added `:74-77` |
| `RangeReference.Evaluate:14` | EXACT | `:14-15` |
| `TableDefinition.cs` (item 5) | GONE | `Table.cs` (Phase 3): record `:26-35`, derived geometry `:39-68`, `TryGetColumnIndex` `:86`, `TryGetColumnRange` `:114-133` (`[#Data]` only, `false` when `DataRowCount == 0` `:121`) |
| `RangeReference.GetBounds:99-104` | MOVED | `:89-105`, min/max `:100-103`; `TryGetBounds` `:112` |
| `ReferenceGuard.cs:87-94` DynamicRange arm; `:96` default; `:3-10`; `:30-36`; `:92-94` | MOVED | DynamicRange `:87-96`; **`default: return null` `:115-116`**; class doc `:3-10` EXACT; method doc `:30-36` EXACT; three Phase 7 arms `:106-113` sit between them |
| `Subtotal.GatherSkippingSubtotals:70` `Sheets[range.SheetName]`; `Subtotal.cs:35-38`; `Subtotal.cs:181-186` | MOVED | file is `Expressions/Mathematics/Subtotal.cs`; guard `:43-46`; the throwing indexer is `AggregateCodes.Gather` `:72` (`:111`, `:140`); the anchored re-dispatch is `AggregateCodes.cs:190-195` |
| `DependencyExtractor.cs:111-116`, `:118`, `:46-48`, `:50-53`, `:216-217`, `:40-44`, `:91-103`, `ResolveName:258-262` | EXACT | all unchanged (`git log`: no engine commit since Phase 1 on this file) |
| `AnchoredFormulaSupport.cs:37`; `:62` | MOVED | `NameReference => true` **`:42`** (rationale `:34-41`); `_ => false` **`:67`** |
| `Workbook.cs:552-555`, `:585-588` (`_namesVersion++`); `:126-133` | GONE | `_definitionsVersion++` at `:609-612`, `:641-644`, `:675-679`; field comment `:155-166`; property `:172` |
| `RecalculationEngine.cs:181-184`; `IsStale:186-210` | GONE | doc `:65-71`; `_definitionsSnapshot` `:87`; sticky `:89-94`; `EnsureFresh` `:208-224`; `HasDefinitionChange` `:230` |
| `ArrayEvaluation.cs:135-136` (Anchored arm), `:146-147` (`Row{[AnchoredRangeReference]}`), `:139-140` cost guard, `:114-117` invariant, `:198-199` default, `:477-479` | MOVED/GONE | Probe `:267-376`: `AnchoredRangeReference` `:282-283`, `OpenRangeReference` `:286-287`, `NameReference` `:293-299`, **`Row { Arguments: [NameReference or Reference] }` `:306-307`** (the anchored-only pattern is gone), default `:373-374`; invariant doc `:135-148` and `:255-257`; `TryBuildOperand` default `:557-559` |
| `ArrayEvaluation.cs:435-437`, `:455-462`, `:232-244` (`ScalarOperand`), `:226-228` (`ArrayOperand.At`) | MOVED | build Anchored `:478-480`; `Row` `:520-527`; `ScalarOperand` **`ArrayOperands.cs:103-115`**; `ArrayOperand` `ArrayOperands.cs:80-100` |
| `NumericAggregation.cs:110` gate; `:107-112`; `:56-101` reference arms in `ArrayEvaluation` | GONE/MOVED | gate is `TryStream` at `:126`; comment `:103-125`; `Probe`'s reference arms `:274-299` |
| `Index.cs:23`, `Index.cs:171`, `OrderStatistics.cs:507` "already write `is not Reference`" | FALSE | `Index.cs:22` `TryStream`; `Index.cs:175-176` `IsBareReferenceNode`; `OrderStatistics.cs` is 414 lines and has no gate (SMALL/LARGE: `OrderSelection.cs:71` `TryStream`) |
| `RangeReference.cs:265-268` (inline dense hit) | EXACT | `:265-271` |
| `Row.cs:13-35`; `:15-28`; `Row.cs:16` | MOVED/EXACT | `:8-42`; typed arms `:15-28` EXACT; `[var only]` `:40`; `TopRow` `:16` EXACT |
| `LookupFunctions.cs:279-297` Column | MOVED | `:271-303`, `[var only]` `:298` |
| `Rows.cs:17-21` `boundOpenRanges:false` | MOVED | `Rows.cs:35-42` via `ReferencePosition.TryResolve` (`:108-114`) |
| `Index.TryResolveReference:171` `is not Reference` | MOVED/FALSE | `:160`; gate `:175-176` is `!IsBareReferenceNode(Arguments[0]) && IsArrayEligible(…)` |
| `NumericAggregation :87-93` Anchored; `:197-203` FoldA; `:126-139` default | EXACT/MOVED | `:87-93` EXACT; FoldA Anchored `:211-217`; default `:140-155` |
| `RangeValueCache.cs:767` `TryGetRangeSnapshot`; gate | EXACT | `:767`, gate `:769`; doc `:757-766` |
| `CriteriaScan.cs:45-47`; `RangeValueCursor.cs:62-65`; `ArgumentFlattening.cs:180-182`; `OrderStatistics.cs:23-25`; `RangeAggregate.Memoize :35` | MOVED/EXACT | `PositionalRange.Open` `:109-115` (gate `:111-113`); `RangeValueCursor.Open` `:59-65` (gate `:63-64`); `ExpandCached` `:204-215` (gate `:209-211`); `Median` `:22-24` EXACT, `Rank` `:127-128`; `Memoize` `RangeValueCache.cs:28-37` (gate `:34-35`) EXACT |
| `DynamicRange.cs:18-19` | EXACT | `:18-19` |
| `ArgumentFlattening.cs:98-101`, `:50-61`, `:109-113`; `RangeValueCursor.cs:93-96`; `PositionalRange.Open` | MOVED | `ExpandComputedValues` anchored `:127-130`, capacity `:134-146`; `FlattenComputedValues` anchored `:67-77`; cursor anchored `:93-96` EXACT; `PositionalRange.Open` anchored `:140-143` |
| `InformationFunctions.cs:113-130`; `LookupFunctions.cs:526-540` | MOVED | `IsFormula` switch `:113-130` EXACT, `#VALUE!` `:135-137`; `FormulaText` `:566-602` |
| `TryResolveReferenceTests.cs` | EXACT | exists, 10 tests |
| `MissingSheetReferenceTests.cs :6-11`; `AssertRef` | EXACT | doc `:6-11`; `Build` `:17`, `AssertRef` `:35`; 58 tests |
| `CaptureValueTests.cs`; `NamedReferences.cs:68`, `:62-68` | MOVED | file exists (2 methods / 4 cases); default at `:79`, switch `:59-80` |
| `MiniCseConsumerTests.cs` "K1 tests" | EXACT | 54 tests; `Sum_OfRowOverLiteralRangeOnMissingSheet_KeepsTheSyntacticGap` `:560` |
| `DependencyExtractorTests.cs :68` `Indirect_IsAlwaysDirty`; "14 today"; `DependencyScan.cs:34` | EXACT/FALSE/GONE | `:68` EXACT; **18** today; `DependencyScan` lives in `DependencyExtractor.cs:28-35` (`IsEmpty` `:34`) |
| `IndirectTests.cs`; `Indirect.cs:32-68`, `:58`, `:68`, `:50-56` | EXACT/MOVED | 7 tests; `:32-69`; `ParseFormulaBody` at `:61`; `:68` EXACT; `:50-56` EXACT |
| `RecalculationEngineTests.cs`; `RecalculationEngine.cs:181-184`; `Workbook.cs:593-606`; `:572` | MOVED | `DefineName` rebuild tests `:253-300`, `DefineTable` twins `:336-396`; `HasUnqualifiedReference` `:797-810` (`_ => false` `:809`); sentinel `:630` |
| `FormulaWriterTests.cs :268` | MOVED | `:272` |
| `RangeValueCache.cs:757-775` | EXACT | `:757-775` |
| `WorksheetStreamLoader.cs:500-515` | MOVED | `:518-528` (Phase 4's audit) |
| `ArrayOperand.At :226-228` | MOVED | `ArrayOperands.cs:89-100` |

## 2. Behavioural claims, checked

True and still load-bearing: `TryResolveRange` takes only a `Workbook` because `DependencyExtractor.Visit` has only `Workbook? wb` (`:78`) — true; `Visit`'s `default: return;` loses the dependency silently (`:216-217`) — true, and `DependencyScan.IsEmpty` (`:34`) still makes a no-workbook scan look like "depends on nothing"; `IsVolatile` must not be overridden because `case Function` (`:207-214`) would mark the formula `AlwaysDirty` — true; `NamedReferences.CaptureValue` needs no arm because its default is `expression.Evaluate` (`:79`) — true, re-measured `LET`/`CHOOSE`/`+` over a name all 60 and the oracle agrees for a table (60 / 60 / 60, both modes); `ComputedValue.EnumerateValues`'s `case Reference` catch-all (`:189-190`) yields the reference value as one element that `AddReferenced` drops — true, so item 3's invariant (a) (return the CONCRETE range, never `Reference(this)`) stands, and `TryBuildScalarConditionIf.WrapScalar` (`:990-998`) is a third reason: it resolves a reference VALUE through `BuildRange` only when it is a `RangeReference`, else a 1x1 `#VALUE!`; `INDEX` and `OFFSET` never call `ReferenceGuard` (M1) — true (`Index.cs`, `Offset.cs`: zero hits); `INDEX(NoSuch,1,1)` = `#REF!`, `OFFSET(NoSuch,1,0)` = `#REF!`, `SUM(NoSuch)` = `#NAME?`, `SUBTOTAL(9,NoSuch)` = `#NAME?` — all re-measured true; `SUBTOTAL` throws `KeyNotFoundException` on a removed sheet without the guard — still structurally true: `AggregateCodes.Gather` indexes `workbook.Sheets[range.SheetName]` (`:72`) and only `Subtotal.cs:43-46`'s guard call protects it; `TryGetRangeSnapshot` rejects anything but `RangeReference`/`OpenRangeReference` (`:769`) while every probe site gates on `argument is Reference` — true at `CriteriaScan.cs:111-113`, `RangeValueCursor.cs:63-64`, `ArgumentFlattening.cs:209-211`, `OrderStatistics.cs:22-24`/`:127-128`, `RangeValueCache.cs:34-35`, so items 15/16 remain PERF-only and correct without the change (a `TableReference` falls to each site's `default:` → `Evaluate` → reference value → `EnumerateValues`); `INDIRECT` works for free through `ParseFormulaBody` (`:61`) + `parsed.TryResolveReference` (`:68`) and needs a cell (`:50-56`) — true, oracle: `SUM(INDIRECT("Tabela1[Valor]"))` 60, `SUM(INDIRECT("Tabela1["&D1&"]"))` 60, `SUM(INDIRECT(E1))` 60 with `E1` = `Tabela1[Valor]`; `HasUnqualifiedReference` reaches `_ => false` for a node with no `SheetName` (`:809`) — true; `DefinedNameArrayEligibilityTests` pins the fifteen shapes — true, 30 cases.

False or stale: "`ROW(T[c])` → `#VALUE!` and `ROW(Rng)` → `#VALUE!` (both measured this session)" — `ROW(Rng)` is 2 now (Phase 1); "`ROWS`'s `TryResolveReference`-false path gives 1.0" — gives the argument's own error; "`SUM/VLOOKUP/ROWS/COLUMNS/SUBTOTAL/COUNTIF/MATCH/XLOOKUP` report `#NAME?`" (M1) — for a name today `VLOOKUP(1,NoSuch,1)` is `#REF!`, `COUNTIF(NoSuch,">0")` is **0**, `MATCH(1,NoSuch,0)` is `#N/A` (oracle: `#NAME?`, `#NAME?`, `#NAME?`; `XLOOKUP(1,NoSuch,NoSuch)` `#N/A` on both) — pre-existing name-class divergences that belong to the sweep, not to this phase, but item 19's matrix must not pin them as table behaviour; "`COUNTA((C1:C3<>"")*1)` => 1" — 3; "`Index.cs:23`, `:171`, `OrderStatistics.cs:507` all already write `is not Reference`" — replaced by `IsBareReferenceNode`/`TryStream`; "`NumericAggregation.cs` is the only `IsArrayEligible` call site lacking the guard" (verification #10) — there is no such call site; "the phase adds no item for S4" (M4) — S4's range half shipped; "a `Tables` dictionary added … deserializes case-SENSITIVE" (M5) — `RestoreComparers` handles it; "`COUNT((T[c]<>"")*1)` → 1 and `MIN(IF(T[c]>0,T[c]))` → 0 with the node absent from `Probe`" — the mechanism still holds (`default:` `(true, false)` broadcasts an opaque scalar) but the numbers now describe a pre-Phase-10 broadcast: on the oracle the targets are 3 and 10 (both **CSE**; PLAIN gives 0 and `#VALUE!`), and `SUM(ROW(Tabela1[Valor]))` is **9 CSE / 2 PLAIN**; "MS documents the specifier as returning null … I have taken to mean `#REF!`" (open question 1) — measured `#REF!` as a VALUE (finding 1); "Neither machine has an oracle" — `/tmp/aspose-probe-fable` answers every question the phase asks.

The design's 23-consumer "free" list is confirmed on the oracle for the shapes it names, with the entry mode: `SUM` 60, `COUNT` 3, `COUNTA` 3, `AVERAGE` 20, `MAX` 30, `ROWS` 3, `COLUMNS` 1, `AREAS` 1, `ISREF` TRUE, `INDEX(…,2,1)` 20, `OFFSET(…,1,0)` `#VALUE!` PLAIN / 20 CSE (a 3x1 result intersected at row 20 — the boundary rule, correct here too), `SUM(OFFSET(…,1,0,2,1))` 50, `VLOOKUP` 20, `LOOKUP` 2, `MATCH` 2, `XLOOKUP` 2, `SUBTOTAL(9)` 60, `AGGREGATE(9,4)` 60, `SUMPRODUCT(T[Valor],T[Qtd])` 140, `SMALL` 10, `LARGE` 30, `MEDIAN` 20, `COUNTIF(…,">15")` 2, `SUMIF` 50, `SUMIFS` 5, `COUNTBLANK` 0, `ISFORMULA` FALSE (value cells), `FORMULATEXT` `#N/A` (no formula in the top-left), `LET`/`CHOOSE`/`+` 60, `SUM(T[Valor]:T[Qtd])` 66, `SUM(Data!B2:INDEX(T[Valor],3,1))` 60, `SUM(FILTER(T[Valor],T[Valor]>15))` 50, `ROWS(FILTER(…))` 2, `SUM(SORT(T[Valor],1,-1))` 60, `INDEX(SORT(…),1)` 30, `ROWS(UNIQUE(T[Item]))` 3 — all both modes unless marked.

## 3. Dependencies

- **Phase 4 T1 — HOLDS, and is the only one.** Nothing on `main` defines `TableReference`, `TableArea` or tag 327; no branch or worktree carries Phase 4. Everything in this phase compiles against T1's node (six-member `TableArea`, `Data = 0`, decoded `TableName`/`ColumnName`, `: Reference`, no `IsVolatile` override, tag 327). Items 8 (T4a), M2 (T2 item 12 + T6a), M3 (T1) are Phase 4's.
- **Phase 3 — RETIRED, with evidence.** `Workbook.Tables` (`:153`), `DefineTable` (`:659`, A1 overload `:690`), `ThrowIfNameTaken` (`:777`), `Table` (`Table.cs:26-35`) with `TryGetColumnRange` (`:114-133`, `[#Data]` only — there is no primitive for `All`/`Headers`/`Totals`/the pairs, so item 5's `TryGetRegion` is still this phase's to write, on `Table`), `DefinitionsVersion` (`:172`) and `RestoreComparers` (`:177-187`).
- **Phase 1 — RETIRED** (item 13, M4, and `ROWS`/`COLUMNS`/`AREAS` error propagation).
- **Phase 11a — RETIRED** (item 12; `IsBareReferenceNode` `:252-253`; the criteria gate `CriteriaScan.cs:234-238`).
- **Phase 7 — merged**; its `ReferenceGuard` arms (`:106-113`), `ProbeIfBranches` (`:397-440`), `TryBuildScalarConditionIf` (`:943-999`), `IArrayProducer` (`ArrayOperands.cs:10-77`) and the producers' `TryBuildSource` all recurse through `Probe`/`TryBuildOperand`, so `FILTER(Tabela1[Valor],…)` works the day the `Probe`/build arms land (oracle 50 / 2 above).
- **Phase 11c — a SCHEDULE dependency, not a design one** (finding 10).
- **Phase 6 — release gate only**: `Danfma.MySheet.Excel` has zero hits for `DefineTable`, `tableParts`, `TableDefinitionPart`, `<table`; no release before Phase 6 (Phase 4's re-verification already says so).

## 4. Blocker and majors

- **B1 — still correct, oracle-confirmed, one number moved** (finding 7).
- **M1 — diagnosis true, prescription inverted** (finding 1). Option (a) (add the guard to `INDEX`/`OFFSET`) would make `INDEX(Tabela1[#Totals],1,1)` short-circuit to the same `#REF!` it already gives, but for an unknown TABLE it would turn `#REF!` into `#NAME?` — the oracle's analogue (`INDEX(NoSuch,1,1)` `#NAME?`) says that IS the right code, yet the same defect exists for every defined name today and the fix is one general rule in the sweep (`Index.cs:28-34`, `Offset.cs:84-88`, `VLookup.cs:20-24`, `MATCH`, the criteria family's error-valued slot), not a table-specific arm. Option (c) with the value-error rule: drop the uniformity claim, drop `tableError` from the guard, pin the actual per-consumer codes with the oracle's `[#Totals]` column beside them and the `NoSuch` analogue for an unknown table.
- **M2 — owned by Phase 4** (T2 item 12: `case TableReference` in `FormulaWriter.Write` + `StructuredReferenceSyntax.Write`; T6a: the export note). Keep this phase's two verifications (round-trip `=SUM(Tabela1[Valor])` through `ToFormula`; export a table-bearing workbook via `SaveAsExcel`) as cross-checks that T2 merged, and note the FORMULATEXT asymmetry disappears with T2.
- **M3 — Phase 4 T1's; numbers stale** (finding 6).
- **M4 — moot** (finding 3); replace with one PIN through `Workbook.GetCellValue`, not `Expression.Evaluate`: `=Tabela1[Valor]` in `Data!E2`/`E3`/`E4` = 10/20/30 and in `Main!H20` = `#VALUE!` (oracle PLAIN column; Phase 4's re-verification measured the in-row values).
- **M5 — moot** (finding 3).

## 5. What the design does not cover and now must

**(a) A `TableReference : Reference` in every Phase 7 / 11a site, decided from the code.** `IsBareReferenceNode` (`:252-253`): true, no change; 11c's overload: true regardless of scope. `TryStream` (`:233`): a bare `SUM(Tabela1[Valor])` takes the reference path — `NumericAggregation.Fold`'s switch has no `TableReference` arm, so it lands in `default:` (`:103-155`), `TryStream` declines, `Evaluate` yields the reference value, `EnumerateValues` expands it (item 14 is the PERF-only struct-enumerator arm, unchanged). `Probe` (`:267-376`): needs the `(true, true)` arm (finding 5); without it `default:` broadcasts an opaque scalar and `COUNT((T[c]<>"")*1)`/`MIN(IF(T[c]>0,T[c]))` go wrong exactly as item 10 says. `TryBuildOperand` (`:463-561`): needs the build arm (finding 5). `ProbeIfBranches` (`:397-440`): a `TableReference` branch is a bare reference — skipped under a scalar condition, probed-but-not-counted under an array condition — and `TryBuildScalarConditionIf` hands a taken table branch through `WrapScalar` → `BuildRange`; that is what the oracle wants: `SUM(IF(TRUE,Tabela1[Valor],0))` 60, `SUM(IF(TRUE,Tabela1[Valor],SEQUENCE(3)))` 60, `ROWS(IF(TRUE,Tabela1[Valor],SEQUENCE(3)))` 3, `SUM(IF(FALSE,Tabela1[Valor],0)*A1:A3)` 0 CSE / `#VALUE!` PLAIN, `COUNTIF(IF(TRUE,Tabela1[Valor],0),">0")` 3 — and the tree already gives 60 / 60 / 3 for a NAME (re-measured), so a table behaves like a name here, NOT like the literal range of sweep item 32 (`SUM(IF(TRUE,Data!B2:B4,0))` `#VALUE!` here, unmoved). `ResolveNameShape` (`:775-799`): unchanged; a defined name whose definition is `Tabela1[Valor]` resolves through `TryResolveRaw` (`:137-159`) → the virtual → `Range`. `ResolvePositionRange` (`:687-735`): resolves the node, re-checks the sheet, degrades to `Scalar` on failure — no arm (finding 5). The criteria gate `RejectComputedArray` (`:234-238`): not rejected (bare reference), `PositionalRange.Open` falls to `ArgumentFlattening.ExpandComputedValues`' `default:` — correct, boxed (item 16). `ReferenceGuard` (`:37-118`): ONE arm, resolved-sheet check only (finding 1). `AggregateCodes.Gather` (`:60-215`): `default:` → `Evaluate` → `TryGetReference` → re-dispatch onto the `RangeReference` arm, nested-skip included; an unresolvable table → error propagates unconditionally (`SUBTOTAL(3,Tabela1[#Totals])` `#REF!` on the oracle). `SelectionProducers.TryBuildSource`: recurses through `Probe`, so a table source is the range operand. `EvaluateCell` (`:365-420`): capture → intersect, no change. `DependencyExtractor.Visit` (`:75-219`): needs item 7's arm, unchanged in shape.

**(b) Bare `=Tabela1`** — finding 9. Add to the tests: `SUM(Tabela1)` 66 / `ROWS` 3 / `COUNTA` 9 / `ISREF` TRUE / `INDEX(Tabela1,2,2)` 20 / `=Tabela1` bare `#VALUE!` PLAIN (2-D), and the DependencyExtractor row for a bare table name. `ExcelGridCellReferenceTests.cs:51-56` stays green under mechanism (a).

**(c) `TryGetRegion` geometry on `Table`**, six areas: `All` `(Left..Right, FirstRow..LastRow)`; `Data` `(FirstDataRow..LastDataRow)`; `Headers` `(FirstRow..FirstRow)`, `HasHeaderRow` false → `Error.Ref`; `Totals` `(LastRow..LastRow)`, `HasTotalsRow` false → `Error.Ref`; `HeadersAndData` `(FirstRow..LastDataRow)`; `DataAndTotals` `(FirstDataRow..LastRow)` — the pairs SHRINK, never error (re-measured on this fixture: `SUM(T[[#Data],[#Totals]])` 66 without and 72 with a totals row, `SUM(T[[#Headers],[#Data]])` 66 both, `ROWS(T[#All])` 4 / 5, `COUNTA(T[#All])` 12 / 14); then column narrowing after the band (`SUM(T[[#Totals],[Valor]])` `#REF!` without a totals row, 0 with one whose Valor total is empty). The `DataRowCount == 0` case is finding 2's ruling; until ruled, `TryGetRegion` must carry a THIRD outcome (empty) rather than folding it into `Error.Ref`, or the ruling cannot be applied without re-opening the primitive.

**(d) Docs** — finding 8; plus a "Structured references" subsection under Tables in `workbook-and-expressions.md` (resolution, error codes with the oracle's column, the empty-body rule once ruled, implicit intersection of a bare `=Tabela1[Valor]`) and its pt-BR twin.

**(e) The error-valued range slot in the criteria family** is a pre-existing divergence this phase will expose for tables: `COUNTIF(NoSuch,">0")` is 0 here and `#NAME?` there, `COUNTIF(1/0,">0")` `#DIV/0!` there; with finding 1's arm, `COUNTIF(Tabela1[#Totals],">0")` will be 0 here and `#REF!` there. Pin it with both numbers and hand it to the sweep (or fix the general rule — one arm in `PositionalRange.Open`'s fallback — if the controller prefers; it is not table-specific).

**(f) Item 17** stands (both `IsFormula` `:113-130` and `FormulaText` `:567-583` fall to `_ => (null, null)` → `#VALUE!` for any node they do not name — re-measured `ISFORMULA(Rng)` `#VALUE!` and `FORMULATEXT(Rng)` `#VALUE!` for a NAME too, a sibling gap for the sweep). The oracle's targets: `ISFORMULA(Tabela1[Valor])` FALSE over value cells, `FORMULATEXT` `#N/A`; unresolvable: `ISFORMULA` FALSE, `FORMULATEXT` `#REF!` — so item 17's arm returns the node's error for `FORMULATEXT` and FALSE, not an error, for `ISFORMULA`.

## 6. Decomposition — five tasks, disjoint files, oracle rows per task

Pre-dispatch rulings (controller): R1 finding 2 (empty data body); R2 finding 1 (the guard returns `null` for an unresolvable table; per-consumer codes are the node's error); R3 finding 9's mechanism for bare `=Tabela1` and its sequencing behind Phase 4 T5; R4 whether (e) is fixed here or in the sweep. Every task rebases onto merged 11c and Phase 4 T1.

- **T1 — Geometry and the resolver's tests** (first, alone). Item 5 on `Table.cs` (`TryGetRegion`, six areas + the empty outcome), `TableReference.TryResolveRange` body if T1 of Phase 4 left it stubbed, `tests/…/TryResolveReferenceTests.cs` (item 18) and `TableTests.cs`. Oracle rows to pin (both modes unless marked): `SUM(T[Valor])` 60; `COUNTA(T[#All])` 12 / 14 with totals; `ROWS(T[#All])` 4 / 5; `COUNTA(T[#Headers])` 3; `SUM(T[[#Headers],[#Data]])` 66; `SUM(T[[#Data],[#Totals]])` 66 / 72; `SUM(T[[#Totals],[Valor]])` `#REF!` / 0; `SUM(T[#Totals])` `#REF!` / 6; `ISREF(T[#Totals])` FALSE; the header-only rows of finding 2 under R1.
- **T2 — Guard and graph.** Item 6 (one arm, `ReferenceGuard.cs`), item 7 (`DependencyExtractor.cs`), items 19, 22 and 24's structured-reference variant (`MissingSheetReferenceTests.cs`, `DependencyExtractorTests.cs`, `RecalculationEngineTests.cs`). Rows: `COUNT(T[#Totals])` 0, `COUNTA(T[#Totals])` 1, `SUM` `#REF!`, `SUBTOTAL(9)`/`(3)` `#REF!`, `ROWS` `#REF!`, `ISREF` FALSE, `COUNT(T[Valor],T[#Totals])` 3, `SUM(Data!B2:B4,T[#Totals])` `#REF!`; the removed-sheet matrix is MySheet's own convention (unmeasurable on the oracle: deleting the sheet deletes the table) and is pinned as such, `SUBTOTAL` included; `DependencyExtractorTests` 18 → 21+.
- **T3 — Mini-CSE arms.** Items 10, 11 (as shrunk), 21 (`ArrayEvaluation.cs`, `MiniCseConsumerTests.cs`), plus the (a)-table pins. Rows: `COUNT((T[Valor]<>"")*1)` **3 CSE** (0 PLAIN); `MIN(IF(T[Valor]>0,T[Valor]))` **10 CSE** (`#VALUE!` PLAIN); `SUM(ROW(T[Valor]))` **9 CSE** (2 PLAIN); `SUM(COLUMN(T[Valor]))` 2; `SUM(T[Valor]*2)` **120 CSE**; `SUM(IF(T[Valor]>15,1,0))` **2 CSE**; `SUM(LEN(T[Item]))` **3 CSE**; `SUM(-T[Valor])` **-60 CSE**; `SUM(T[Valor]*A1:A3)` **320 CSE** with `A1:A3` = 5,0,9; `SUM((T[Valor]>15)*T[Qtd])` **5 CSE**; `SUMPRODUCT((T[Valor]>15)*1)` 2; the corpus idiom `SMALL((ROW(T[c])-ROW(INDEX(T[c],1,1))+1)/((T[c]<>"")*(T[c]<>0)),1)` 1 and its `AGGREGATE(15,6,…,1)` twin 1; `SUM((T[#Totals]<>"")*1)` `#REF!`, `COUNT((T[#Totals]<>"")*1)` 0, `COUNTA(…)` 1, `SUM(T[#Totals]*A1:A3)` `#REF!`, `COUNT(T[#Totals]*A1:A3)` 0, `SUM(IF(T[#Totals]>0,1,0))` `#REF!`; the IF rows of §5(a); `COUNTIF(T[Valor]*1,">0")` **`#REF!` CSE** (`#VALUE!` PLAIN); `SUM(FILTER(T[Valor],T[Valor]>15))` 50, `ROWS` 2, `SUM(SORT(…,1,-1))` 60, `ROWS(UNIQUE(T[Item]))` 3.
- **T4 — Consumers and bindings** (beside T3, no shared file). Items 14, 15, 16, 17 (`NumericAggregation.cs`, `RangeValueCache.cs`, `ArgumentFlattening.cs`, `RangeValueCursor.cs`, `CriteriaScan.cs`, `InformationFunctions.cs`, `LookupFunctions.cs` FORMULATEXT only) and items 20, 23 (`CaptureValueTests.cs`, `IndirectTests.cs`) plus the M4 boundary pin. Rows: the 23-consumer list in §2 with its two PLAIN/CSE splits; `LET`/`CHOOSE`/`+` 60; `INDIRECT` 60 ×3; `ISFORMULA(T[Valor])` FALSE, `FORMULATEXT` `#N/A`, `ISFORMULA(T[#Totals])` FALSE, `FORMULATEXT(T[#Totals])` `#REF!`; the (e) divergence pinned with both numbers unless R4 fixes it; `RangeValueCacheEquivalenceTests` green after item 15.
- **T5 — Bare table name and docs** (after Phase 4 T5 merges; last). Finding 9 under R3 (`Parser.cs` or `NamedReferences.cs`/`NameReference.cs` per the mechanism), the docs of finding 8 with the twin diff, `ExcelGridCellReferenceTests.cs`/`ExpressionParserTests.cs` as needed. Rows: `SUM(Tabela1)` 66, `ROWS` 3, `COLUMNS` 3, `COUNTA` 9, `ISREF` TRUE, `INDEX(Tabela1,2,2)` 20, `COUNTIF(Tabela1,">15")` 2, `LET(t,Tabela1,SUM(t))` 66, `SUM(INDIRECT("Tabela1"))` 66, `COUNT((Tabela1<>"")*1)` **9 CSE**, `=Tabela1` bare `#VALUE!` PLAIN.

Verification plan corrections: #2 expects **1927 + new** on a tree where 11c has closed the red pin (0 failures), else 1 known; #5 `DependencyExtractorTests` **18 + 3**, and the item is 22, not 23; #9 is `grep -c '^\[MemoryPackUnion' … = 328` and belongs to T1 of Phase 4; #10 is deleted (no `IsArrayEligible` call sites to check — `TryStream`/`IsBareReferenceNode` is the guard, pinned by `DefinedNameArrayEligibilityTests`' 30 cases); add: `dotnet build --no-incremental` before any `--no-build` run; every oracle row above is a formula result with an assertion behind it, never a comment.

Risks retired: #1 (finding 4), #5 (finding 5), #6 (item 12 shipped; the scan-order argument is now `TryStream`'s doc `:205-212` and the 2-D name pin). Risks kept: #3 (composite canonicalization, consistent with Phase 4 item 12/16 — Aspose keeps `[[#Data],[Valor]]` as written, MySheet will normalize it to `[Valor]`, a harmless writer divergence to note), #4 (the `MissingSheet` name is even more honest after finding 1 — the arm checks a sheet and nothing else), #5's replacement: a future edit that makes the resolved sheet check return `tableError` re-introduces the COUNT/COUNTA divergence silently, so T2's pins for `COUNT(T[#Totals])` 0 and `COUNTA` 1 are the guard.

Open questions, current state: Q1 answered (finding 1); Q2 (unknown table `#NAME?` vs `#REF!`) unmeasurable on the oracle — it rejects the formula at entry — and the design's `#NAME?` stands by analogy with the measured `NoSuch` column; Q3 delivered by 11a; Q4 answered against the design (finding 2, needs R1); Q5 answered and moved here by Ruling 3 (finding 9, needs R3).
