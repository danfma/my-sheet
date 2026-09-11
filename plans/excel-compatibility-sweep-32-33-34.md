# Sweep items 32, 33, 34 + the deferred minors

Close the three deliberate divergences the structured-table epic recorded, plus the two test-quality minors deferred by the Phase 4 fix wave. Plan of record for this epic; the sweep file (`phase-11-excel-compatibility-sweep.md` items 31-34) holds the measured numbers each item was pinned with.

## For Future Agents
As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to `Complete` and write its **Phase Summary**; run the phase's **Verification Plan** and record the result. Rules that bind: the suite is never red at a merge (both suites `failed: 0` via `dotnet run --project tests/<proj>/<proj>.csproj -c Release`, NEVER `dotnet test`); no AI attribution in any commit; one worktree per concurrent task; "Excel" means Aspose.Cells 26.6.0 as MEASURED (copy `/tmp/aspose-probe-fable` before editing; name the entry mode of every number); every merge is **ff-only after a clean rebase**; TDD red-first, and **never silently change an expected value** — where a recorded divergence pin flips, the commit body carries BOTH numbers and the reason. Controller reviews each task before the next dispatch; a multi-part final review (controller subagent + Copilot, findings verified before acting) closes the branch, then the fix wave, rebase, ff-merge.

**Design decisions made by the controller up front** (so no implementer has to guess):
- Item 33 is attacked via a real empty-reference representation (Ruling R1's option a), NOT by special-casing `#REF!` consumers. The shape: `TableReference.TryResolveRange` maps `TableRegionOutcome.Empty` to a zero-row reference value instead of `Error.Ref`; every consumer that pattern-matches `RangeReference` gets the oracle's row for the empty case, measured from the two tables already recorded (sweep item 33 + the Phase 5 re-verification's finding 2, both modes) and re-measured where missing. `Absent` (the `[#Totals]`-with-no-totals-row singleton) STAYS `#REF!` — the oracle still errors there.
- Item 34's two classes get ONE rule each, not per-consumer arms: (a) an error-valued argument in the criteria family's range slot propagates (one general arm in `PositionalRange.Open`'s fallback, exactly as the sweep line prescribes); (b) a resolving consumer must return the argument's OWN error when the node it resolved from is an unresolvable name — the `#NAME?` the name node already carries — instead of the consumer's fallback code. R2's COUNT/COUNTA pins (0 / 1) MUST survive: they are different functions on a different path, and any fix that turns them into `#REF!` is wrong (their pins are the guard).
- Item 32 covers BOTH recorded shapes: the scalar-condition `IF` with a bare-reference branch (the four pinned rows) and the reference-valued `CHOOSE`/`IF` branch under an operator (`SUM(CHOOSE(1,T[Valor])*2)` `#VALUE!` vs 120 CSE — the literal range answers identically, so the fix is generic to the branch type, not table-specific).

## Phase 1: Item 34 — error propagation in the criteria family and the resolving consumers
Status: Not started

Worktree `/Volumes/Work/Develop/MySheet-33`, branch `feat/sweep-32-33-34` off `main` @ `b778ab7`. Files per the sweep: `CriteriaScan.cs` (`PositionalRange.Open`), the criteria family consumers, the resolving consumers (`Index`, `VLookup`, `Match`/`LookupFunctions`), `MiniCseConsumerTests.cs`, `MissingSheetReferenceTests.cs`, docs twins.

- [ ] Task 1 — red pins for BOTH classes from the recorded numbers (`COUNTIF(NoSuch,">0")` `#NAME?`, `COUNTIF(1/0,">0")` `#DIV/0!`, `COUNTIF(Tabela1[#Totals],">0")`/`SUMIF`/`COUNTBLANK` `#REF!`; `VLOOKUP`/`INDEX` `#NAME?` and `MATCH` `#NAME?` over an unknown name AND an unknown table — re-measure the whole resolving family on the oracle first, both modes, because the recorded rows cover VLOOKUP/INDEX/MATCH but XLOOKUP/HLOOKUP/OFFSET and kin were not all recorded), then the two rules, then the rewrite of the divergence pins that close (`AResolvingConsumer_OverAnUnknownTable_ReportsItsOwnCode_ADivergence`, the criteria-zero pins) with both numbers in the commit bodies.
- [ ] Controller review + R2 guard check: `COUNT(Tabela1[#Totals])` 0 and `COUNTA` 1 untouched (mutation: the new criteria arm must not reach them).
- [ ] Docs: every sentence that recorded these divergences as open flips to closed, both twins.

### Verification Plan
- Filtered red-first runs recorded; then full gates: core and Excel `failed: 0` (counts recorded in the ledger), csharpier clean, Release 0 warnings.
- R2 mutation: revert the criteria arm → the new pins red; COUNT/COUNTA pins green throughout.

### Phase Summary
_(write when the phase completes)_

## Phase 2: Item 32 — the bare-reference branch under a scalar condition and under an operator
Status: Not started

Files per the sweep: `If.cs`, `ArrayEvaluation.cs` (`ProbeIfBranches`, `TryBuildScalarConditionIf`), the CHOOSE/mini-CSE arms, `CriteriaScan.cs`, `MiniCseConsumerTests.cs`, docs twins. The four pinned rows (`SUM(IF(TRUE,A1:A3,0))` 14, `SUM(IF(TRUE,A1:A3,SEQUENCE(3)))` 14, `ROWS(...)` 3, the single-cell name 0) turn red when fixed — their values live in `ABareReferenceBranch_UnderAScalarConditionIf_IsUnmovedAndStillDiverges`. The CHOOSE seam (`SUM(CHOOSE(1,Tabela1[Valor])*2)` 120 CSE vs `#VALUE!`, literal range identical) is the operator half; measure the IF-twin (`SUM(IF(TRUE,T[Valor],0)*2)` or equivalent) on the oracle before fixing so the rule covers the family.

- [ ] Task 2 — re-measure the family on the oracle (both modes, scalar-condition IF × {bare reference, producer, mixed, single-cell name} × {bare, under operator}), red pins, the fix (stream/capture the bare-reference branch like the producer path; make the reference-valued branch lift under an operator like a syntactic range), the pinned-divergence rewrites.
- [ ] Controller review + regression check: Phase 11c's binding-site pins (LET/CHOOSE/`+` stay ranges at a top level) and the criteria-gate pins (`COUNTIF(+A1:A3,">0")` 2) must stay green — the fix must not widen `IsBareReferenceNode`'s top-level semantics.

### Verification Plan
- As Phase 1's, plus: the 11c pins (`DefinedNameArrayEligibilityTests`, `ArrayBindingTests`) green unchanged.

### Phase Summary
_(write when the phase completes)_

## Phase 3: Item 33 — the empty-reference representation
Status: Not started

The biggest item: a zero-row reference value that behaves like the oracle's empty reference (`SUM`/`COUNT`/`COUNTA` 0, `ROWS` 0, `COLUMNS` 1, `AREAS` 1, `ISREF` TRUE, `AVERAGE` `#DIV/0!`, `MIN`/`MAX` 0, `SMALL`/`MEDIAN` `#NUM!`, `INDEX` `#REF!`, `MATCH`/`VLOOKUP` `#N/A`, `COUNTIF`/`SUMIF`/`COUNTBLANK` 0, `SUBTOTAL` 0, `SUMPRODUCT` 0, `OFFSET` 0, `FILTER` `#CALC!`, `ROW` 2-plain/1-CSE, bare `=T[Valor]` `#VALUE!`-plain/0-CSE, `ISERROR` TRUE-plain/FALSE-CSE, `LET`/`IF`/`INDIRECT` 0) while `Absent` stays `#REF!`. The measured tables: sweep item 33 + the Phase 5 re-verification finding 2 (header-only column, both modes). Wire shape must round-trip (a table that loads with zero data rows from a real `.xlsx` — the committed `f7-header-only` fixture is exactly this — must resolve, not `#REF!`).

- [ ] Task 3 — design + red pins + the representation + the per-consumer sweep, guided by the two recorded tables and re-measured where a consumer is missing from them.
- [ ] Controller review: `Absent` still `#REF!`; the f7 fixture's registration pins untouched; a MemoryPack/wire round-trip of the new shape if it adds a node or field.

### Verification Plan
- As Phase 1's, plus: the Excel suite must stay green with the f7 fixture LOADING (its evaluation pins flip from `#REF!` to the oracle's numbers — both numbers in the commit bodies).

### Phase Summary
_(write when the phase completes)_

## Phase 4: the deferred minors + the sweep file closes
Status: Not started

- [ ] The two Phase 4 fix-wave deferrals, controller-executed: assert the warning on the decode tests that use the options-less overload where the overload allows it, and turn `Load_SharedStructuredReferenceMaster`'s coinciding "42" cache into a 999-style lie so the assertion can never pass by reading the cache.
- [ ] Sweep file: items 32, 33, 34 marked CLOSED-BY with the phase names and the commit hashes (controller-owned).
- [ ] Docs: any survivor that still calls these three divergences open (grep "recorded divergence", "deliberate", the Tables blockquote's zero-data-rows row) flips to the new truth, both twins, parity-checked.

### Verification Plan
- Both suites `failed: 0`, csharpier clean, Release 0 warnings, twin parity counts equal.

### Phase Summary
_(write when the phase completes)_

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete — same as the epic's: rebase the branch onto `main`, `git merge --ff-only`, all gates on `main`, then the next release absorbs it)_
