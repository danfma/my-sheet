# Sweep items 31, 35-43 + the Aspose.Cells 26.7 oracle migration

Close the Excel-compatibility divergences a downstream consumer's corpus still hits — Bug 6 `INDEX(range,0,n)` (item 37), Bug 8's `MATCH` residual, and error literals in formula text (item 43) — plus the ordinary-range gaps sweep 32-34 surfaced (items 35-42) and item 31, and migrate the oracle to Aspose.Cells 26.7.0 with an audit of every older recorded matrix. Source of the divergences: `~/MYSHEET-CALC-DIVERGENCES.md` (consumer-owned) and the draft items in `.superpowers/sdd/sweep-32-33-34/phase4-sweep-edits.draft.md`.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to `Complete` and write its **Phase Summary** (what was done, key decisions, anything needed to continue with zero context); run the phase's **Verification Plan** and record the result before moving on. When all phases are done, fill in **Final Recap** and **Deployment Plan**.

Rules that bind every phase:
- Both suites must show `failed: 0` at every merge. Run them with `dotnet run --project tests/<proj>/<proj>.csproj -c Release`, NEVER `dotnet test`.
- No AI attribution in any commit.
- One worktree per concurrent task.
- TDD red-first. **Never silently change an expected value**: a flipped pin carries BOTH numbers and the reason.
- Every merge is ff-only after a clean rebase.
- **"Excel" means Aspose.Cells 26.7.0 as MEASURED** (user decision, 2026-09-14).
  - Copy a probe before editing it; `/tmp/aspose-probe-fable` is shared and never edited.
  - One formula per workbook: the oracle can depend on evaluation order.
  - Name the entry mode (PLAIN / CSE) of every number.
  - Record 26.6.0 beside 26.7.0 for every new row.
- **Review (user decision, 2026-09-14; replaces the earlier Fable review loop to save tokens):**
  - The controller verifies gates only.
  - Each phase's deferred review goes into `.superpowers/sdd/pending-verifications.md`.
  - After all phases, one Fable or GLM (zclaude) session works through that list.
  - Releasing with some inconsistency is accepted, with fixes later.
- Ledger: `.superpowers/sdd/sweep-31-35-43/progress.md`. Briefs sit beside it.

**Design decisions made by the controller up front:**
- **Branch and parallelism.** Branch `feat/sweep-31-35-43`, worktree `/Volumes/Work/Develop/MySheet-43`, off `feat/sweep-32-33-34` @ `4b82423`.
  - It rebases onto `main` once sweep 32-34 merges.
  - Phases 0, 1 and 5 start immediately: they are measurement, or parser code that sweep 32-34 does not touch.
  - Phases 2-4 write code only after that rebase, because sweep 32-34's pre-merge work is still moving `LookupFunctions.cs` / `ReferencePosition.cs`.
- **Oracle version.** Aspose.Cells 26.7.0. A 26.6/26.7 difference in an OLD pin is an audit finding (Phase 5) with both numbers, never an automatic flip.
- **Item 43** reuses the existing `ErrorValue` AST node, which is already a MemoryPack union member, so no new union tag is expected. If one proves necessary, allocate the next free tag under the append-only rule and justify it.
- **Item 37:** `INDEX` with row or column 0 returns a REFERENCE (the area's column / row), so reference consumers see a range and not a computed array. Measure on the oracle which consumers treat it as a reference BEFORE designing the arms.
- **Acceptance.** The consumer's shapes are acceptance rows. At the end, the divergence probe (`.superpowers/sdd/sweep-32-33-34/divergence-probe/`, bumped to 26.7.0) shows no DIFF row for Bug 6, Bug 8, the Bug 9 remainder (item 31) or error literals. Rows that follow the array-entered column by design, `-0`, and the circular self-count stay as documented; they are out of scope.

## Phase 0: Oracle measurement on Aspose 26.7 (no code)
Status: Complete

- [x] Copy the shared probe to `/private/tmp/claude-501/sweep-31-35-43/phase0/` and bump it to Aspose.Cells 26.7.0 (keep a 26.6.0 twin).
- [x] Measure every row of items 31, 35, 36, 38, 39, 40, 41, 42 and the Bug 8 residual, PLAIN and CSE, 26.6.0 beside 26.7.0.
  - Bug 8 rows: the value slot over an error cell for MATCH (exact and approximate), XMATCH, XLOOKUP, VLOOKUP, HLOOKUP and LOOKUP.
  - Item 42 at several formula-cell rows, since implicit intersection depends on the row.
- [x] Record the binding table in `.superpowers/sdd/sweep-31-35-43/phase0-oracle.md`.

### Verification Plan
- `phase0-oracle.md` has a 26.6.0 / 26.7.0 × PLAIN / CSE row for every formula named in the draft items 31 and 35-42 and in the Bug 8 list, and names every row where the two versions differ.
- **Result:** 115 rows across all nine groups. The 26.6.0 and 26.7.0 columns are identical on every row. The controller cross-checked five rows against independent earlier measurements, and all agree.

### Phase Summary
- 115 rows were measured, and 26.6.0 = 26.7.0 on every one, so Phases 3-4 need no version-specific handling.
- Bug 8 residual: sweep 32-34 closed it before merging (`5975c63`: a lookup value that is a single cell holding an error leads MATCH / XMATCH / XLOOKUP). The only thing left in the value slot is the range-in-the-value-slot lifting divergence (`SUM(MATCH(A5:A7,A5:A7,0))`, oracle `#VALUE!` / 6).
- **Item 36:** the oracle contradicts itself. Its `ROWS` / `COLUMNS` over `OFFSET` ignore an explicit height/width while `SUM` / `COUNT` honour it. Ruling: MySheet uses one coherent window (explicit size when given, else the base's), and the size-ignoring `ROWS` / `COLUMNS` rows are registered as an oracle defect.
- **Item 42:** the oracle uses the same rule for literal ranges and table columns (PLAIN is row-position implicit intersection, CSE takes the first element). Ruling: MySheet's table reference converges on the literal-range convention MySheet already has.
- Both rulings are recorded in the ledger's "Phase 0 — ACCEPTED" entry.

## Phase 1: Item 43 — error literals in formula text
Status: Complete

- [x] Measure on 26.7.0 (26.6.0 beside), PLAIN and CSE:
  - which error literals parse and what they evaluate to: `#REF!`, `#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`, `#CALC!`, `#GETTING_DATA`, lowercase forms;
  - error literals inside functions and operators;
  - the qualifier forms `Sheet1!#REF!`, `'My Sheet'!#REF!` and `A1:#REF!`.
- [x] Red pins in the tokenizer/parser tests, the evaluation tests, and the xlsx loader tests (a formula containing `#REF!` must load, not degrade through `UnparsableFormula`).
- [x] Tokenizer and parser produce `ErrorValue` for an error literal; the qualifier forms follow the measurement.
- [x] Formula-text round trip: `FORMULATEXT`, export/save writes the literal back, and serialization is unchanged (union count stays 328).
- [x] Both docs twins record the new truth, checked sentence by sentence.
- [x] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- Divergence probe rows `=#REF!`, `=SUM(#REF!)`, `=IF(ISNA(#N/A),1,0)`, `MATCH(#REF!,#REF!,0)`, `MATCH(1,#REF!,0)`, `IFNA(MATCH(#REF!,#REF!,0),"na")` and `COUNTIF(#REF!,#REF!)` give MySheet = Aspose PLAIN.
- `dotnet csharpier check .` is clean, and `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` shows 0 warnings.
- Both suites `failed: 0`, and `grep -c '^\[MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs` is 328.

### Phase Summary
- **What changed.** The tokenizer and parser now read error literals as `ErrorValue` in formula text. That covers the plain literals, the literals inside functions and operators, and the deleted-reference continuations (`Sheet1!#REF!`, `#REF!A1`, `#REF!Tabela1`, grid-bounded). The xlsx loader no longer degrades those formulas through `UnparsableFormula`.
- **Commits.** `5c7eeab`, `678dc65`, `bc753f9`, `79cb19d`, `82b9de0` on `feat/sweep-31-35-43`.
- **Review.** The opencode review loop took 5 rounds to reach Clean (`.superpowers/sdd/sweep-31-35-43/review/phase-1-round-5.md`). Gates at the close: core 3042/0, Excel 137/0, union count 328, 0 warnings.
- **Registered, not fixed.**
  - Formula-text write-back drops a sheet qualifier on `#REF!` (Phase 1 review I-2).
  - Error criteria in the COUNTIF family (M-5).

  Both are listed in `.superpowers/sdd/pending-verifications.md`.

## Phase 2: Item 37 (Bug 6) — `INDEX` with row or column 0
Status: Complete

Scope widened by user decision (2026-09-14), because the corpus idiom `INDEX($5:$1000,0,MATCH(h,$4:$4,0))` inside arithmetic needs two more root causes fixed:
- (a) whole-row / whole-column positional coordinates are absolute, not relative to the populated box (`MATCH("x",$4:$4,0)` 1 vs 3);
- (b) `INDEX`/`OFFSET` references lift under an operator the way item 32 made `CHOOSE`/`IF` lift (`SUMPRODUCT((INDEX(E5:H10,0,1)>6)*1)` 1 vs 4).

- [x] Rebase the branch onto `main` after sweep 32-34 merges; gates green.
- [x] Measure on 26.7.0 `INDEX(area,0,n)`, `INDEX(area,n,0)` and `INDEX(area,0,0)` through every consumer:
  - aggregates and counts: SUM, COUNT, COUNTA, SUMPRODUCT, the COUNTIF-family range slot, and AGGREGATE in both forms;
  - reference readers: ROWS, COLUMNS, ROW, COLUMN, AREAS, ISREF;
  - lookups and nesting: the MATCH lookup array, nested `INDEX`, and the OFFSET base;
  - bare cell (implicit intersection), and a header-only table band;
  - the corpus shape `IFERROR(AGGREGATE(15,6,(ROW(INDEX(r,0,MATCH(…)))-ROW(INDEX(INDEX(r,0,MATCH(…)),1,1))+1)/((INDEX(r,0,MATCH(…))<>"")*(…)),ROWS($B$2:B2)),"")`.
- [x] Red pins, then `INDEX` returns the row/column reference, and each consumer arm follows the measurement.
- [x] The empty-table row `SUM(INDEX(Tabela1[Valor],0,1))` flips with both numbers.
- [x] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- Divergence probe rows 37-44, 72, 73 and 86 match the oracle.
- Mutation: reverting the zero-argument arm turns only the new pins red.
- Gates as in Phase 1.

### Phase Summary
- **Item 37: zero-axis INDEX.** `INDEX(area,0,n)`, `INDEX(area,n,0)` and `INDEX(area,0,0)` return a row, column or area reference, consumed by every reference-aware consumer (`d522027`, `0315cee`).
- **(a) Whole-row and whole-column positional coordinates are absolute** (`81f228e`). MATCH and XMATCH over an open range translate through retained source positions. Open-range exact MATCH uses the shared snapshot and exact index; over 500k cells it went from 11.8 ms to 0.034 ms per evaluation.
- **(b) INDEX and OFFSET references lift under operators** (`e04ae14`, `38bd0b7`).
- **Review loop fixes**, reviewed by sol, sol, then luna, Clean at round 3: `a21ff92`, `29453d0`, `e4e4e57`.
- **Adversarial pass over the deferred `0315cee` checklist.** Reviewed by sol and terra (x2), then luna, Clean at round 3: `a0372df`, `42c5527`, `3c5297b`, `9ff0951`. It fixed:
  - `INDEX(Ghost!A1:A3,0,1)` now gives `#REF!` (was `#VALUE!`);
  - a negative INDEX index now gives `#VALUE!` (was `#REF!`);
  - two docs overclaims.
- **Integrated** on `feat/sweep-31-35-43` @ `9ff0951`: core 3069/0, Excel 137/0, union count 328. All Phase 2 acceptance rows in the divergence probe match. Rows 41 and 43, first registered as gaps, were fixed by (a) and (b).
- **Registered, not fixed:**
  - the HLOOKUP row index past a one-row table: MySheet `#REF!` vs Aspose `#N/A`, where the oracle is inconsistent with its own single-cell form;
  - the INDEX `area_num` fourth argument is unsupported;
  - array constants `{…}` do not parse.

## Phase 3: The lookup family — Bug 8 residual + items 38, 40, 41
Status: Complete. The review loop ran 5 rounds, then per the user's decision the branch was integrated and its residuals were fixed on the integration branch (`07b43d0`..`37b5335`), with a closure review.

- [x] Bug 8 residual: whatever sweep 32-34's pre-merge MATCH fix leaves open in the value slot over an error cell, across MATCH, XMATCH, XLOOKUP, VLOOKUP, HLOOKUP and LOOKUP.
- [x] Item 38: XLOOKUP checks that its lookup and return arrays agree in size.
- [x] Item 40: XLOOKUP's array slots over an unresolvable argument (the per-slot oracle split).
- [x] Item 41: a resolving consumer over a non-reference argument (`VLOOKUP(1,5,1)`, a single-cell name holding an error, `VLOOKUP(1,1/0,1)`).
- [x] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- The Phase 0 table rows for these items all match MySheet.
- Divergence probe rows 56-62 match.
- Gates as in Phase 1.

### Phase Summary
- **Item 38 — XLOOKUP shapes.**
  - XLOOKUP validates the lookup axis. A 1x1 lookup array is a ROW. This was measured after the controller's unmeasured "both axes" ruling proved wrong. A 2D lookup array is `#VALUE!`, and the return array must match along the lookup axis.
  - The result is the matched row or column, for both reference and computed return arrays: `SUM(XLOOKUP(2,A1:A3,B1:C3))` is 220 and `INDEX(XLOOKUP(2,SEQUENCE(3),SEQUENCE(3,3)),1,2)` is 5. The earlier parallel row-major cursors, which returned 100 instead of 20, had been wrong since before this sweep.
  - Operands are bound once, and LET/computed-name arrays are accepted.
  - A syntactically reference-returning XLOOKUP is read as that reference by the shared readers (criteria, AGGREGATE/SUBTOTAL, OFFSET) through one helper, `NamedReferences.TryResolveReferenceReturningNode`.
  - A miss propagates `#N/A` or the `if_not_found` value. A non-reference fallback in a reference slot is `#REF!` or `#VALUE!`.
  - A LET-bound XLOOKUP is a value array: reference slots refuse it, and `SUM` reads the whole row.
- **Item 40.** Unresolvable array slots follow the oracle per slot.
- **Item 41.**
  - Resolving consumers over a non-reference follow the CSE column.
  - Single-cell tables go through the same 1x1 validation, including the negative-index `#VALUE!` that Phase 2 introduced.
  - Direct error-cell tables keep the measured precedence (`#N/A` exact, `#DIV/0!` approximate), and an error lookup value propagates over a single-cell table.
- **Bug 8 residual.** Sweep 32-34's I1/M1/M2 were folded in: VLOOKUP/HLOOKUP over a 1x1 error range give `#DIV/0!`, the named-error IF branch is pinned, and an XML-doc overclaim was removed.
- **Process.**
  - Review loop: sol for r1-r5, plus one stalled review killed and redispatched without subagents.
  - Fix rounds: sol for r1-r5b.
  - Integration: `integration-p3` (sol), which resolved an `Offset.cs` production conflict in Phase 4's favour.
  - Integration fix rounds: r1-r3, with a sol closure review.
  - Branch head `2c20292`, integrated as `07b43d0`.
- **Registered, not fixed:** sweep items 51-56 and the XLOOKUP `ROWS`/`COLUMNS` oracle defects.

## Phase 4: Items 35, 36, 39, 42 and 31
Status: Complete (review loop closed at round 6: 5 rounds plus one extra round approved by the user; integrated into `feat/sweep-31-35-43` at `71abb67`)

- [x] Item 35: `SUMIF` / `AVERAGEIF` resize the sum range to the criteria range's shape.
- [x] Item 36: `OFFSET` with height/width omitted inherits the base's size.
- [x] Item 39: `ROWS` / `COLUMNS` (and siblings) of a `LET`-bound reference.
- [x] Item 42: one ISERROR / N / IFERROR convention for literal ranges and table references, decided by the Phase 0 measurement in both modes.
- [x] Item 31: the criteria gate for an array-conditioned `IF` returning references, plus the MIXED selector residual from sweep 32-34.
- [x] Flip the empty-table rows in `EmptyTableReferenceTests.TheRowsThatDependOnAnOrdinaryRangeGap_KeepTheEnginesAnswer` with both numbers.
- [x] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- The Phase 0 table rows for these items match MySheet.
- Divergence probe row 71 (item 31) matches the CSE column.
- The 11c fence (`DefinedNameArrayEligibilityTests`, `ArrayBindingTests`) stays unchanged.
- Gates as in Phase 1.

### Phase Summary
- **Item 35.** `SUMIF`/`AVERAGEIF` resize the sum range from its top-left corner to the criteria shape. Strict `SUMIFS`/`AVERAGEIFS` shape mismatches return `#VALUE!`.
- **Item 36.** `OFFSET` with an omitted height or width inherits the base size. Negative sizes extend the window up or left, zero sizes give `#REF!`, and both corners are validated against the grid.
  - Ruling: MySheet keeps one coherent window. Aspose's `ROWS`/`COLUMNS` over an explicit size (including `ROWS(OFFSET(A:A,…))`) is a registered oracle defect.
- **Item 39.** A LET-bound reference stays a reference for `ROWS`, `COLUMNS`, `ROW`, `AREAS`, `ISREF` and value consumers. A LET-bound computed array in a criteria range slot is still refused.
- **Item 42.** The whole scalar-consumer family routes through `ScalarReferenceValue` with implicit intersection, for literal ranges and table references alike: `ISERROR`, `ISERR`, `ISNA`, `N`, `IFERROR`, `IFNA`, `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISBLANK`. A failed reference resolution is not evaluated twice.
  - Ruling: a formula inside its own range keeps MySheet's `#REF!` cycle guard.
- **Item 31.** Array-conditioned `IF` reference selectors are admitted in criteria and sum-range slots. One classifier, `PositionalRange.TryResolveSelectorReference` with `SelectorRoute` values `Structural`, `ElementWise` and `NotASelector`, resolves selectors through `NamedReferences.TryResolveReference`.
  - Value slots validate the final selected reference, so a missing sheet gives `#REF!`.
  - Criteria slots give `#REF!` only for structural routes (LET/name chains, and INDEX, OFFSET or INDIRECT that end in a reference). IF/CHOOSE routes follow the CSE column (0).
- **Process.**
  - Review loop: rounds r1-r5 with sol, plus a user-approved r6 with terra.
  - Fix rounds r1-r6 with sol; pin rounds r1-r2 with terra, run in a parallel worktree with disjoint file ownership.
  - Branch head `fe09af4`, cherry-picked onto the integration line as `71abb67` (core 3185/0, Excel 137/0).
- **Registered, not fixed:** three pre-existing resolve-then-evaluate double-draw sites (`ArrayEvaluation` IF branch, `LookupFunctions` CHOOSE, `ReferencePosition`).
- **Routed elsewhere:** the XLOOKUP-miss-in-criteria finding went to Phase 3 and was fixed in `2c20292`.

## Phase 5: Aspose 26.7 audit of the older recorded matrices (no code)
Status: Complete (stopped by user decision 2026-09-14; the remaining coverage is registered)

- [x] Inventory every recorded oracle matrix:
  - the ledgers and probe logs under `.superpowers/sdd/phase-*/`;
  - the probe directories under `/private/tmp/claude-501/`, excluding the sweep-32-33-34 task probes, which sweep 32-34 re-measures itself;
  - `/tmp/aspose-probe-fable`;
  - the measured rows of the sweep file's items;
  - test comments that cite oracle numbers.
- [x] Re-run what can be re-run on 26.7.0 and diff it against the recorded 26.6.0 numbers.
- [x] Every difference becomes a new sweep item with both numbers and its pinning test. No pin flips.

### Verification Plan
- `.superpowers/sdd/sweep-31-35-43/oracle-26.7-audit.md` states coverage (matrices found / re-run / not re-runnable, with the reason) and lists every difference with file:line of its pin.

### Phase Summary
- About 1,870 distinct formulas were re-run on both Aspose.Cells 26.6.0 and 26.7.0 with **0 differences**. They came from the test-comment rows, every phase ledger, and about 150 ledger-only review rows; see `.superpowers/sdd/sweep-31-35-43/oracle-26.7-audit.md`. No new sweep item and no pin flip resulted.
- The user stopped the audit on 2026-09-14 to save tokens. Still open: about 127 of the 130 probe files in two old scratchpads, which were never re-run. That gap is registered in `.superpowers/sdd/pending-verifications.md`.

## Phase 6: Docs, sweep file, lessons, final review, merge
Status: Complete. The Fable 5.1 gate and its fix wave are merged into `main` at `40a3f8f` (2026-09-15). Not pushed; the user owns the push.

- [x] Docs twins, sentence-level parity for every behaviour changed (each phase review checked this). One oracle-version note per doc twin records the 26.7.0 migration and the zero-difference audit. The per-sentence 26.6.0 citations stay, because not every sentence was re-measured (terra `docs-oracle-note-brief.md`).
- [ ] Sweep file: items 31 and 35-43 marked CLOSED-BY with hashes; the Phase 5 audit items added.
- [x] Lessons appended to `tasks/lessons.md`: 15 from this sweep, plus the model-routing correction.
- [x] Final divergence-probe run, branch vs `main`, recorded in the ledger. On Aspose 26.7.0, 141 rows: 48 changed, 46 toward CSE and 0 away; the other 2 throw everywhere. Every remaining non-CSE row is registered or ruled. The whole-branch quality review is the Fable gate on 2026-09-15.
- [x] Rebase onto `main`, gates on `main`, local ff-only merge (`main` `b8d9968` → `5b51ffc`, via the integration branch `feat/sweep-31-integrate`). No push; the user owns the release.
- [x] **Final Fable 5.1 code-quality verification (user requirement, 2026-09-14):** done 2026-09-15; see the Phase Summary.
  - When: after EVERY branch of this correction effort is merged into `main` (sweep 32-34 is already merged; this branch, and any follow-up correction branch from the `~/MYSHEET-CALC-DIVERGENCES.md` list), and BEFORE any push.
  - Scope: the whole delta `98b330b..main`.
  - A Claude Fable 5.1 reviewer judges code quality: design, duplication, correctness risks, test quality, docs.
  - Its Critical/Important findings get a fix wave (merged into `main`) before the user pushes.

### Verification Plan
- Both suites `failed: 0` on `main` after the merge.
- csharpier clean, Release 0 warnings.
- Twin parity checked sentence by sentence.
- The divergence probe shows no DIFF row in the acceptance groups.

### Phase Summary
- **2026-09-14 — integration and local merge.**
  - Integration order: Phases 1 and 2 (including the adversarial fixes) first, then Phase 4 (`71abb67`), then Phase 3 (`07b43d0`; the `Offset.cs` conflict was resolved in Phase 4's favour).
  - Integration fix rounds r1-r3 (`0edca24`..`3f066ae`), then a docs oracle note (`5b51ffc`).
  - Final probe against the old `main`: 46 rows moved toward CSE and 0 away.
  - Local ff merge, followed by the plans/lessons commit `e1a3270`.
- **2026-09-15 — Fable 5.1 final gate** (`.superpowers/sdd/sweep-31-35-43/final-gate/fable.md`) over `98b330b..main`: 1 Critical, 3 Important, 8 Minor.
  - C1: XLOOKUP was evaluated 2-9× per consumer, which gave silently wrong values under volatile arguments.
  - I1: MATCH resolved its lookup array 3×.
  - I2: SUMIF/COUNTIF/AVERAGEIF built the snapshot on first read.
  - I3 (pre-existing): open-range XLOOKUP mapped the populated ordinal to a row.
  - Minors: VLOOKUP/HLOOKUP duplication, dead code, an allocation, a docs overclaim, stale comments, style.
  - Item 57 was measured and refuted.
- **Fix wave.**
  - Part A (sol): `3dd0dcf` one memo per XLOOKUP node, `489b592`, `5e4848d` open-range XLOOKUP selected by cell coordinate (oracle-measured 30/50/50/9), `41d9ca3`.
  - Part B (terra): `9d7419d` second-use snapshot admission restored, `da5417d`, `9d324a9` shared `LookupTable`.
  - The combined closure review was Clean.
  - A controller benchmark over DISTINCT formula cells then showed the wave had regressed open-range XMATCH ~50,000× and XLOOKUP ~2.3×. The review's timing was a cache hit.
  - Part C (sol): `a1b01d9` added the uncached `--open-range-lookup` benchmark mode; `40a3f8f` restored the snapshot/exact-index paths.
  - Result vs the pre-gate `main`: XLOOKUP 8.2 → 0.18 ms, SUM/COUNTIF over XLOOKUP 16.5/8.0 → 0.003/0.14 ms, MATCH/XMATCH ≤ 0.001 ms. An independent benchmark closure review was Clean.
- **Final `main` `40a3f8f`.** Core 3392/0, Excel 137/0, Release 0 warnings, union 328, attribution 0. The divergence probe is byte-identical to `e1a3270`.
## Final Recap
Sweep 31/35-43 closed sweep items 31 and 35-43, migrated the Excel oracle to Aspose.Cells 26.7.0, and passed a whole-range code-quality gate.
- **Oracle.** Phase 0 measured 115 rows and the Phase 5 audit about 1,870 formulas, with 0 differences between 26.6.0 and 26.7.0.
- **Fixes by item.**
  - Item 43 (Phase 1): error literals in formula text.
  - Item 37 (Phase 2): zero-axis INDEX, absolute open-range coordinates, INDEX/OFFSET lifting.
  - Items 38/40/41 (Phase 3): XLOOKUP shape, axis and reference semantics; resolving consumers over non-references.
  - Items 35/36/39/42/31 (Phase 4): SUMIF resize, OFFSET windows, LET-bound references, the scalar-consumer family, criteria selectors behind one route classifier.
- **Process.** Every phase ran implement → review loops on opencode (GPT-5.6 sol/terra, bounded by explicit rules in `.superpowers/sdd/AGENT-RULES.md`) against the measured oracle, with extra rounds approved by the user where a loop hit its limit.
  - Registered, not fixed: items 44-56 and the ROWS/COLUMNS oracle defects (phase-11 sweep file).
  - The final Claude Fable 5.1 review found real evaluation-count and performance defects that the loops had missed. Both were fixed and benchmark-verified with uncached measurements.
  - Lessons are in `tasks/lessons.md`: unmeasured controller rulings, fixture-bound expected values, subagent stalls, cache-hit benchmarks.

## Deployment Plan
The user owns the release; agents never push.
1. `git fetch origin` and confirm `git log --oneline origin/main..main` lists the sweep 32-34 + 31/35-43 commits and that `origin/main` has not moved (`git merge-base --is-ancestor origin/main main`).
2. Optional re-check on the release machine: `dotnet csharpier check .`, `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` (0 warnings), `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` and the Excel suite (both `failed: 0`), and `dotnet run -c Release --project benchmarks/Danfma.MySheet.Benchmark -- --open-range-lookup` (XLOOKUP under 1 ms, MATCH/XMATCH at or under 0.001 ms).
3. `git push origin main`.
4. Follow-ups stay in the phase-11 sweep file (items 44-56) and `.superpowers/sdd/pending-verifications.md`.
