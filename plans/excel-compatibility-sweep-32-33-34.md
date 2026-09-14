# Sweep items 32, 33, 34 + the deferred minors

Close the three deliberate divergences the structured-table epic recorded, plus the two test-quality minors deferred by the Phase 4 fix wave. Plan of record for this epic; the sweep file (`phase-11-excel-compatibility-sweep.md` items 31-34) holds the measured numbers each item was pinned with.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to `Complete` and write its **Phase Summary**; run the phase's **Verification Plan** and record the result.

Rules that bind:
- The suite is never red at a merge: both suites `failed: 0` via `dotnet run --project tests/<proj>/<proj>.csproj -c Release`, NEVER `dotnet test`.
- No AI attribution in any commit.
- One worktree per concurrent task.
- "Excel" means Aspose.Cells 26.6.0 as MEASURED. Copy `/tmp/aspose-probe-fable` before editing, and name the entry mode of every number. (Superseded on 2026-09-14 by the user's move to 26.7.0; this epic was re-measured on 26.7.0 before its merge, see the Final Recap.)
- Every merge is **ff-only after a clean rebase**.
- TDD red-first, and **never silently change an expected value**: where a recorded divergence pin flips, the commit body carries BOTH numbers and the reason.
- The controller reviews each task before the next dispatch. A multi-part final review (controller subagent + Copilot, findings verified before acting) closes the branch, followed by the fix wave, the rebase and the ff-merge.

**Design decisions made by the controller up front** (so no implementer has to guess):
- **Item 33** is attacked via a real empty-reference representation (Ruling R1's option a), NOT by special-casing `#REF!` consumers.
  - The shape: `TableReference.TryResolveRange` maps `TableRegionOutcome.Empty` to a zero-row reference value instead of `Error.Ref`.
  - Every consumer that pattern-matches `RangeReference` gets the oracle's row for the empty case. Those rows come from the two tables already recorded (sweep item 33 + the Phase 5 re-verification's finding 2, both modes), re-measured where missing.
  - `Absent` (the `[#Totals]`-with-no-totals-row singleton) STAYS `#REF!`, because the oracle still errors there.
- **Item 34**'s two classes get ONE rule each, not per-consumer arms:
  - (a) An error-valued argument in the criteria family's range slot propagates. One general arm in `PositionalRange.Open`'s fallback, exactly as the sweep line prescribes.
  - (b) A resolving consumer returns the argument's OWN error when the node it resolved from is an unresolvable name (the `#NAME?` the name node already carries), instead of the consumer's fallback code.
  - R2's COUNT/COUNTA pins (0 / 1) MUST survive. They are different functions on a different path, and any fix that turns them into `#REF!` is wrong; their pins are the guard.
- **Item 32** covers BOTH recorded shapes:
  - the scalar-condition `IF` with a bare-reference branch (the four pinned rows);
  - the reference-valued `CHOOSE`/`IF` branch under an operator (`SUM(CHOOSE(1,T[Valor])*2)`: `#VALUE!` here vs 120 CSE). The literal range answers identically, so the fix is generic to the branch type, not table-specific.

## Phase 1: Item 34 — error propagation in the criteria family and the resolving consumers
Status: Complete

Worktree `/Volumes/Work/Develop/MySheet-33`, branch `feat/sweep-32-33-34`, cut from `main` @ `98b330b`. Files per the sweep:
- `CriteriaScan.cs` (`PositionalRange.Open`) and the criteria family consumers;
- the resolving consumers (`Index`, `VLookup`, `Match`/`LookupFunctions`);
- `MiniCseConsumerTests.cs`, `MissingSheetReferenceTests.cs`;
- the docs twins.

- [x] Task 1 — red pins for BOTH classes from the recorded numbers, then the two rules, then the rewrite of the divergence pins that close (`AResolvingConsumer_OverAnUnknownTable_ReportsItsOwnCode_ADivergence`, the criteria-zero pins), with both numbers in the commit bodies.
  - Criteria class: `COUNTIF(NoSuch,">0")` `#NAME?`, `COUNTIF(1/0,">0")` `#DIV/0!`, `COUNTIF(Tabela1[#Totals],">0")`/`SUMIF`/`COUNTBLANK` `#REF!`.
  - Resolving class: `VLOOKUP`/`INDEX` `#NAME?` and `MATCH` `#NAME?` over an unknown name AND an unknown table.
  - Re-measure the whole resolving family on the oracle first, both modes: the recorded rows cover VLOOKUP/INDEX/MATCH, but XLOOKUP/HLOOKUP/OFFSET and kin were not all recorded.
- [x] Controller review + R2 guard check: `COUNT(Tabela1[#Totals])` 0 and `COUNTA` 1 untouched (mutation: the new criteria arm must not reach them).
- [x] Docs: every sentence that recorded these divergences as open flips to closed, both twins.

### Verification Plan
- Filtered red-first runs recorded; then full gates: core and Excel `failed: 0` (counts recorded in the ledger), csharpier clean, Release 0 warnings.
- R2 mutation: revert the criteria arm → the new pins go red; COUNT/COUNTA pins stay green throughout.
- **Result:** red first 5/1/2; core 2552 → 2561 / 0, Excel 133 / 0. Mutation forcing `IsOwnSlotError` false turned exactly the new pins red, with COUNT 0 / COUNTA 1 intact. Reproduced independently by the final review (3 red).

### Phase Summary
- **Rule (a)** (`a542f56`, pre-rebase `558f4c8`): `PositionalRange.Open`'s fallback plus the `RangeValueCursor` mirror propagate an argument whose OWN value is an error, guarded by `IsOwnSlotError`. That guard keeps reference nodes, resolving names and the collapse artifact on their old paths.
- **Rule (b):** `ReferencePosition.TryUnresolvedError`, routed through VLOOKUP, HLOOKUP, INDEX, MATCH, XMATCH, OFFSET, LOOKUP and FORMULATEXT.
- **Brief claim measured false:** MATCH over `[#Totals]` answered `#N/A`, not `#REF!`. It now answers the node's `#REF!`.
- **Residuals:** XLOOKUP's array slots (sweep item 40) and a resolving consumer over a non-reference argument (item 41).
- **Pre-merge regression.** A corpus-derived divergence probe found that rule (b) had turned `MATCH(A1,A1)` over an error cell from `#DIV/0!` into `#N/A`. `5975c63` fixed it with `ReferencePosition.IsLookupValueError`: a single cell holding an error leads MATCH, XMATCH and XLOOKUP. The same fix closed a downstream consumer's `MATCH(A1,A1,0)` divergence ("Bug 8").
- **Final review:** `4ecd5b5` restored MATCH's approximate-path error return (a range-collapse lookup value had started answering a position) and counts a 1x1 range holding an error as a single cell.
- **Docs:** no item-34 sentence existed in the twins.

## Phase 2: Item 32 — the bare-reference branch under a scalar condition and under an operator
Status: Complete

Files per the sweep: `If.cs`, `ArrayEvaluation.cs` (`ProbeIfBranches`, `TryBuildScalarConditionIf`), the CHOOSE/mini-CSE arms, `CriteriaScan.cs`, `MiniCseConsumerTests.cs`, docs twins.
- The four pinned rows turn red when fixed; their values live in `ABareReferenceBranch_UnderAScalarConditionIf_IsUnmovedAndStillDiverges`:
  - `SUM(IF(TRUE,A1:A3,0))` 14;
  - `SUM(IF(TRUE,A1:A3,SEQUENCE(3)))` 14;
  - `ROWS(...)` 3;
  - the single-cell name, 0.
- The CHOOSE seam (`SUM(CHOOSE(1,Tabela1[Valor])*2)`: 120 CSE vs `#VALUE!` here, literal range identical) is the operator half. Measure the IF twin (`SUM(IF(TRUE,T[Valor],0)*2)` or equivalent) on the oracle before fixing, so the rule covers the family.

- [x] Task 2 — re-measure the family on the oracle (both modes; scalar-condition IF × {bare reference, producer, mixed, single-cell name} × {bare, under operator}), red pins, the fix, and the pinned-divergence rewrites. The fix: stream/capture the bare-reference branch like the producer path, and make the reference-valued branch lift under an operator like a syntactic range.
- [x] Controller review + regression check: Phase 11c's binding-site pins (LET/CHOOSE/`+` stay ranges at a top level) and the criteria-gate pins (`COUNTIF(+A1:A3,">0")` 2) must stay green. The fix must not widen `IsBareReferenceNode`'s top-level semantics.

### Verification Plan
- As Phase 1's, plus: the 11c pins (`DefinedNameArrayEligibilityTests`, `ArrayBindingTests`) green unchanged.
- **Result:** core 2562 / 0, Excel 133 / 0. M1 mutation (the new predicate arms off) → exactly 4 red; M2 (branch counting off) → exactly 2 red; the 11c suites stayed green throughout. The final review reproduced both counts.

### Phase Summary
- `a8ef927` (pre-rebase `c6fde1f`), one rule in four arms: a scalar-condition `IF`/`CHOOSE` whose branches are all bare references carries the taken branch's reference.
  1. `IsBareReferenceNode` If/Choose arms;
  2. `ProbeBranches` counts the streamable branch;
  3. `TryBuildScalarConditionIf` streams it through `BranchValue`;
  4. `If`/`Choose` carry the resolved reference.
- **Closed:** all four recorded rows, the CHOOSE seam (120 CSE), and `COUNTIF(IF(TRUE,A1:A3,B1:B3),">0")` 2, in both modes.
- **Residual:** the MIXED selector at the criteria gate (a computed sibling keeps the node array-eligible) is item 31's, recorded on item 31.
- **Final-review regression, fixed:** the single-cell shape handed scalar consumers a reference (`IF(A1>0,B1,C1)*2` gave `#VALUE!`). `77aea32` returns the cell's value and lets `NumericAggregation` resolve the reference, so `SUM(IF(TRUE,MyCell,0))` stays 0. A volatile LET-bound condition drawn twice was fixed in `f3c09a1`.
- **Docs:** the twins flipped. A stale pt-BR sentence left inside the rewritten paragraph was found by a sentence-level audit and fixed in `dae37b8`.

## Phase 3: Item 33 — the empty-reference representation
Status: Complete

The biggest item: a zero-row reference value that behaves like the oracle's empty reference, while `Absent` stays `#REF!`. Target behaviour:
- aggregates and counts: `SUM`/`COUNT`/`COUNTA` 0, `MIN`/`MAX` 0, `AVERAGE` `#DIV/0!`, `SMALL`/`MEDIAN` `#NUM!`, `COUNTIF`/`SUMIF`/`COUNTBLANK` 0, `SUBTOTAL` 0, `SUMPRODUCT` 0;
- geometry: `ROWS` 0, `COLUMNS` 1, `AREAS` 1, `ISREF` TRUE, `ROW` 2 plain / 1 CSE;
- lookups and selection: `INDEX` `#REF!`, `MATCH`/`VLOOKUP` `#N/A`, `OFFSET` 0, `FILTER` `#CALC!`;
- cell-level: bare `=T[Valor]` `#VALUE!` plain / 0 CSE, `ISERROR` TRUE plain / FALSE CSE, `LET`/`IF`/`INDIRECT` 0.

The measured tables: sweep item 33 + the Phase 5 re-verification finding 2 (header-only column, both modes). The wire shape must round-trip: a table that loads with zero data rows from a real `.xlsx` must resolve, not `#REF!`. The committed `f7-header-only` fixture is exactly this.

- [x] Task 3 — design + red pins + the representation + the per-consumer sweep, guided by the two recorded tables and re-measured where a consumer is missing from them.
- [x] Controller review: `Absent` still `#REF!`; the f7 fixture's registration pins untouched; a MemoryPack/wire round-trip of the new shape if it adds a node or field.

### Verification Plan
- As Phase 1's, plus: the Excel suite must stay green with the f7 fixture LOADING (its evaluation pins flip from `#REF!` to the oracle's numbers, both numbers in the commit bodies).
- **Result:** core 2766 → 2802 / 0 after the fix wave, Excel 134 / 0, union count 328, f7 load/save/load pinned. Mutations: the Empty arm back to `#REF!` → 178 red, all empty-reference pins; the empty-stream arm off → 22 red.

### Phase Summary
- **Representation** (`0acb66a`, pre-rebase `df8b3a1`): a runtime-only `EmptyRangeReference` (zero rows, anchored at the row after the header), produced only by `TableReference.TryResolve`, with one arm per consumer. It is not a union member: count 328, no golden change.
- **The oracle is evaluation-order dependent** over a header-only table (primed vs cold). Controller ruling:
  - the PRIMED core column binds;
  - a classification rule applies: match what a zero-row reference at the anchor can produce; register as an oracle defect anything that needs the empty reference to yield a cell;
  - five pre-existing ordinary-range gaps are pinned with both numbers and registered as sweep items 35-39;
  - ISERROR/N/IFERROR follow MySheet's existing table-reference convention.
- **Independent review** found two host-visible crashes (FILTER/UNIQUE on the columns axis over an empty band), a SUMPRODUCT zero-row shape check, a missing recalc test and four nits. The fix wave (`e998533`, `09b8a0a`) resolved them all.
  - Its design is oracle-backed: a selection keeps its zero rows, and `SelectionProducers.Select` maps only an EMPTY selection to `#CALC!`.
  - It adds one rectangle-only table resolution helper and one bounds helper.

## Phase 4: the deferred minors + the sweep file closes
Status: Complete

- [x] The two Phase 4 fix-wave deferrals, controller-executed:
  - assert the warning on the decode tests that use the options-less overload, where the overload allows it;
  - turn `Load_SharedStructuredReferenceMaster`'s coinciding "42" cache into a 999-style lie so the assertion can never pass by reading the cache.
- [x] Sweep file: items 32, 33, 34 marked CLOSED-BY with the phase names and the commit hashes (controller-owned).
- [x] Docs: any survivor that still calls these three divergences open (grep "recorded divergence", "deliberate", the Tables blockquote's zero-data-rows row) flips to the new truth, both twins, parity-checked.

### Verification Plan
- Both suites `failed: 0`, csharpier clean, Release 0 warnings, twin parity counts equal.
- **Result:** gates green on the final head `557304d`: core 2900 / 0, Excel 134 / 0, 0 warnings. Twin parity was checked sentence by sentence, not only by counts. The one survivor (a stale pt-BR sentence) is fixed.

### Phase Summary
- **Minors** (`99d71e4`):
  - The two unparsable-formula decode tests (cached string / cached error) now load through the options overload and assert the `UnparsableFormula` warning. Mutation: 10 red.
  - `WithoutOptions_DoesNotThrow` and the overflow test keep the options-less overload, since it is their subject or contract.
  - The shared-master cache is 999. The plan's wording was inverted for that test: the lie proves the value came FROM the cache.
- **Sweep file:**
  - items 32, 33 and 34 are CLOSED-BY with the rebased hashes;
  - item 31 carries the mixed-selector contact note;
  - items 35-43 are registered and owned by `plans/excel-compatibility-sweep-31-35-43.md`: the ordinary-range gaps, INDEX row/column 0, the XLOOKUP / VLOOKUP residuals, the ISERROR reference-kind inconsistency, and error literals.
- **Lessons:** appended to `tasks/lessons.md`.

## Final Recap
All three recorded divergences are closed, along with the deferred minors and a downstream consumer's `MATCH` error-propagation divergence.

**Commits** (rebased on `main` @ `a02ed5d`):
- `a542f56` item 34;
- `a8ef927` item 32;
- `0acb66a` item 33;
- `99d71e4` minors;
- `dae37b8` pt-BR docs fix;
- `e998533` + `09b8a0a` item 33 fix wave;
- `5975c63` the MATCH regression fix and Bug 8;
- `77aea32` C1, `4ecd5b5` I1 + I4, `f3c09a1` I3, `20dd2e1` I2 pins, `557304d` minors (the final-review fix wave).

**Suites:** core 2552 → 2900 / 0, Excel 133 → 134 / 0.

**What changed in the engine:**
- **Error propagation:** the criteria family's range slot and the resolving consumers now propagate the argument's own error.
- **Selectors:** a scalar-condition `IF`/`CHOOSE` over bare references carries its reference.
- **Empty tables:** a header-only table's structured references evaluate as an empty reference instead of `#REF!`, crash-free across the selection producers.
- **Lookups:** a single cell holding an error leads MATCH, XMATCH and XLOOKUP.

**Found along the way and registered:**
- the oracle's evaluation-order dependence over header-only tables;
- the item 36 / item 42 oracle findings;
- eight follow-up items (35-42) plus error literals (43), now owned by the next epic.

**Oracle migration.** The user moved the oracle to Aspose.Cells 26.7.0 before the merge. Every row this epic recorded (T1 320, T2 100, T3 + fix wave 1,455, divergence probe 91) was re-measured on 26.7.0 with ZERO differences.

**Final review:**
- Copilot CLI (`gpt-5.4`): "Mergeable: Yes", no findings, mutations and 26.7.0 spot checks reproduced.
- Fable 5.1: "Mergeable: No".
  - Critical C1: item 32's single-cell branch handed scalar consumers a reference, so `IF(A1>0,B1,C1)*2` gave `#VALUE!` instead of 2.
  - Also I1-I4 and seven minors.
  - The controller reproduced every blocking finding with the divergence probe, and the fix wave fixed or pinned them.
  - A re-review is deferred to `.superpowers/sdd/pending-verifications.md`, after the user dropped the review loop to save tokens.

## Deployment Plan
1. On the branch worktree: confirm `feat/sweep-32-33-34` is rebased on the current `main` (`git -C /Volumes/Work/Develop/MySheet-33 rebase main` if `main` moved), then run the gates:
   - `dotnet csharpier check .`
   - `dotnet build Danfma.MySheet.slnx -c Release --no-incremental` (0 warnings)
   - both suites via `dotnet run --project tests/<proj>/<proj>.csproj -c Release`
2. From the main worktree `/Volumes/Work/Develop/MySheet`: `git merge --ff-only feat/sweep-32-33-34`.
3. On `main`, run the same gates (expected core 2900 / 0, Excel 134 / 0).
4. No push by the controller. The user owns the 3.21.0 release (CHANGELOG, tag, publish).
5. Cleanup: `git worktree remove /Volumes/Work/Develop/MySheet-33-rv-fable` and `…-rv-copilot`.
6. The next epic's branch `feat/sweep-31-35-43` was cut from the pre-rebase `4b82423`. Move it with `git -C /Volumes/Work/Develop/MySheet-43 rebase --onto main 4b82423`, only while no agent is working in that worktree.
