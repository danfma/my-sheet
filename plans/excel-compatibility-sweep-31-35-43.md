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
Status: In progress

- [ ] Measure on 26.7.0 (26.6.0 beside), PLAIN and CSE:
  - which error literals parse and what they evaluate to: `#REF!`, `#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`, `#CALC!`, `#GETTING_DATA`, lowercase forms;
  - error literals inside functions and operators;
  - the qualifier forms `Sheet1!#REF!`, `'My Sheet'!#REF!` and `A1:#REF!`.
- [ ] Red pins in the tokenizer/parser tests, the evaluation tests, and the xlsx loader tests (a formula containing `#REF!` must load, not degrade through `UnparsableFormula`).
- [ ] Tokenizer and parser produce `ErrorValue` for an error literal; the qualifier forms follow the measurement.
- [ ] Formula-text round trip: `FORMULATEXT`, export/save writes the literal back, and serialization is unchanged (union count stays 328).
- [ ] Both docs twins record the new truth, checked sentence by sentence.
- [ ] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- Divergence probe rows `=#REF!`, `=SUM(#REF!)`, `=IF(ISNA(#N/A),1,0)`, `MATCH(#REF!,#REF!,0)`, `MATCH(1,#REF!,0)`, `IFNA(MATCH(#REF!,#REF!,0),"na")` and `COUNTIF(#REF!,#REF!)` give MySheet = Aspose PLAIN.
- `dotnet csharpier check .` is clean, and `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` shows 0 warnings.
- Both suites `failed: 0`, and `grep -c '^\[MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs` is 328.

### Phase Summary
_(write when phase completes)_

## Phase 2: Item 37 (Bug 6) — `INDEX` with row or column 0
Status: Not started (blocked on the sweep 32-34 merge)

- [ ] Rebase the branch onto `main` after sweep 32-34 merges; gates green.
- [ ] Measure on 26.7.0 `INDEX(area,0,n)`, `INDEX(area,n,0)` and `INDEX(area,0,0)` through every consumer:
  - aggregates and counts: SUM, COUNT, COUNTA, SUMPRODUCT, the COUNTIF-family range slot, and AGGREGATE in both forms;
  - reference readers: ROWS, COLUMNS, ROW, COLUMN, AREAS, ISREF;
  - lookups and nesting: the MATCH lookup array, nested `INDEX`, and the OFFSET base;
  - bare cell (implicit intersection), and a header-only table band;
  - the corpus shape `IFERROR(AGGREGATE(15,6,(ROW(INDEX(r,0,MATCH(…)))-ROW(INDEX(INDEX(r,0,MATCH(…)),1,1))+1)/((INDEX(r,0,MATCH(…))<>"")*(…)),ROWS($B$2:B2)),"")`.
- [ ] Red pins, then `INDEX` returns the row/column reference, and each consumer arm follows the measurement.
- [ ] The empty-table row `SUM(INDEX(Tabela1[Valor],0,1))` flips with both numbers.
- [ ] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- Divergence probe rows 37-44, 72, 73 and 86 match the oracle.
- Mutation: reverting the zero-argument arm turns only the new pins red.
- Gates as in Phase 1.

### Phase Summary
_(write when phase completes)_

## Phase 3: The lookup family — Bug 8 residual + items 38, 40, 41
Status: Not started (blocked on the sweep 32-34 merge)

- [ ] Bug 8 residual: whatever sweep 32-34's pre-merge MATCH fix leaves open in the value slot over an error cell, across MATCH, XMATCH, XLOOKUP, VLOOKUP, HLOOKUP and LOOKUP.
- [ ] Item 38: XLOOKUP checks that its lookup and return arrays agree in size.
- [ ] Item 40: XLOOKUP's array slots over an unresolvable argument (the per-slot oracle split).
- [ ] Item 41: a resolving consumer over a non-reference argument (`VLOOKUP(1,5,1)`, a single-cell name holding an error, `VLOOKUP(1,1/0,1)`).
- [ ] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- The Phase 0 table rows for these items all match MySheet.
- Divergence probe rows 56-62 match.
- Gates as in Phase 1.

### Phase Summary
_(write when phase completes)_

## Phase 4: Items 35, 36, 39, 42 and 31
Status: Not started (blocked on the sweep 32-34 merge)

- [ ] Item 35: `SUMIF` / `AVERAGEIF` resize the sum range to the criteria range's shape.
- [ ] Item 36: `OFFSET` with height/width omitted inherits the base's size.
- [ ] Item 39: `ROWS` / `COLUMNS` (and siblings) of a `LET`-bound reference.
- [ ] Item 42: one ISERROR / N / IFERROR convention for literal ranges and table references, decided by the Phase 0 measurement in both modes.
- [ ] Item 31: the criteria gate for an array-conditioned `IF` returning references, plus the MIXED selector residual from sweep 32-34.
- [ ] Flip the empty-table rows in `EmptyTableReferenceTests.TheRowsThatDependOnAnOrdinaryRangeGap_KeepTheEnginesAnswer` with both numbers.
- [ ] Append the phase's deferred review checklist to `.superpowers/sdd/pending-verifications.md`.

### Verification Plan
- The Phase 0 table rows for these items match MySheet.
- Divergence probe row 71 (item 31) matches the CSE column.
- The 11c fence (`DefinedNameArrayEligibilityTests`, `ArrayBindingTests`) stays unchanged.
- Gates as in Phase 1.

### Phase Summary
_(write when phase completes)_

## Phase 5: Aspose 26.7 audit of the older recorded matrices (no code)
Status: In progress

- [ ] Inventory every recorded oracle matrix:
  - the ledgers and probe logs under `.superpowers/sdd/phase-*/`;
  - the probe directories under `/private/tmp/claude-501/`, excluding the sweep-32-33-34 task probes, which sweep 32-34 re-measures itself;
  - `/tmp/aspose-probe-fable`;
  - the measured rows of the sweep file's items;
  - test comments that cite oracle numbers.
- [ ] Re-run what can be re-run on 26.7.0 and diff it against the recorded 26.6.0 numbers.
- [ ] Every difference becomes a new sweep item with both numbers and its pinning test. No pin flips.

### Verification Plan
- `.superpowers/sdd/sweep-31-35-43/oracle-26.7-audit.md` states coverage (matrices found / re-run / not re-runnable, with the reason) and lists every difference with file:line of its pin.

### Phase Summary
_(write when phase completes)_

## Phase 6: Docs, sweep file, lessons, final review, merge
Status: Not started

- [ ] Docs twins, sentence-level parity for every behaviour changed. The docs' oracle-version sentences say Aspose.Cells 26.7.0.
- [ ] Sweep file: items 31 and 35-43 marked CLOSED-BY with hashes; the Phase 5 audit items added.
- [ ] Lessons appended to `tasks/lessons.md`.
- [ ] Final divergence-probe run, branch vs `main`, recorded in the ledger (the whole-branch review is deferred to `pending-verifications.md`).
- [ ] Rebase onto `main`, gates on `main`, local ff-only merge. No push; the user owns the release.

### Verification Plan
- Both suites `failed: 0` on `main` after the merge.
- csharpier clean, Release 0 warnings.
- Twin parity checked sentence by sentence.
- The divergence probe shows no DIFF row in the acceptance groups.

### Phase Summary
_(write when phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete)_
