# Release 3.20.0 and the table chain — execution plan

Goal: turn `main`'s one deliberately red test green (Phase 11c), cut release 3.20.0 on a fully green suite, then deliver the structured-table chain (Phases 4, 5, 6) for 3.21.0. This file sequences the four phase files; it does not repeat their items.

## For Future Agents

As work proceeds: mark checkboxes `- [x]` as items complete; when a phase is done, set its status to `Complete` and write its **Phase Summary** (what was done, key decisions, anything needed to continue with zero context); run the phase's **Verification Plan** and record the result before moving on. When all phases are done, fill in **Final Recap** and **Deployment Plan**.

Rules that bind every phase here, decided by the user on 2026-09-10 and 2026-09-11:

- **The suite is never red at a merge.** `dotnet run --project tests/Danfma.MySheet.Tests/Danfma.MySheet.Tests.csproj -c Release` and the Excel twin must report `failed: 0` on `main` before any release is cut. Never `dotnet test`.
- **Cadence per phase:** subagent TDD tasks from the generated briefs, per-task review, then a TWO-part final review — Fable (subagent, model `fable`) and GLM-5.3 (the z.ai endpoint; export the seven env vars from `/Users/danfma/bin/zclaude` and call `$HOME/.local/bin/claude --print … < brief`, never the wrapper, which drops its flags). Copilot is out. Then a fix wave, `git rebase main`, `git merge --ff-only`.
- **No AI attribution in commit messages.** The harness supplies trailers in a system reminder; the user's global rule forbids them and `main` is clean.
- **"Excel" means Aspose.Cells 26.6.0 as measured**, array-entered column. A divergence is a work item.
- Ledgers and briefs live under `.superpowers/sdd/<phase>/` (gitignored); each ledger has a RESUME HERE section.
- **Model per task (user, 2026-09-11): Opus for the complex ones, Sonnet only where the work is mechanical.** Opus: anything that measures a new oracle rule, inverts a design decision, touches `ArrayEvaluation.cs`/`Parser.cs`, or writes an engine arm — Phase 4 T2 and T5, Phase 5 T1, T2, T3 and T5, Phase 11c T3. Sonnet: shifting anchors, pinning numbers another task already measured, and doc sweeps — Phase 4 T6, Phase 5 T4, Phase 6 T2 and T3, Phase 11c T4. A doc sweep on Sonnet still carries the twin-diff requirement: after any scripted edit, grep for a distinctive token from every paragraph you did not mean to touch.
- **Fable ran out of usage credits on 2026-09-11** and killed three subagents mid-flight (nothing lost — every worktree was clean at its last accepted commit). The two-part final review was planned as Fable plus GLM-5.3; until Fable's credits reset, the second opinion is an Opus subagent, since GLM runs on the separate z.ai endpoint and is unaffected.

## Phase A: 11c — array bindings, suite to zero failures

Status: Complete

Plan: `structured-table-references-and-aggregate/phase-11c-array-bindings.md`. Briefs: `.superpowers/sdd/phase-11c-array-bindings/task-{1..4}-brief.md`. Worktree `/Volumes/Work/Develop/MySheet-11c`, branch `feat/array-bindings`.

- [x] Task 1 — the pins, red on arrival (items 1-3), plus the twelve unmeasured oracle rows. `6ad7cea`, core 1927/1 → 1978/38.
- [x] Task 2 — the scope carries an array; gates consult it (items 4-8). `1fd687d`, core → 1988/15 (the 15 are Task 3's).
- [x] Task 3 — defined names, CHOOSE, unary `+` (items 9-12). `ceb7539`, core 1988/15 → **2001 / 0**.
- [x] Task 4 — docs, design of record, bookkeeping (items 13-15). `199fb98`.
- [x] Two-part final review (an Opus subagent and GLM-5.3; both Yes with fixes, same Critical found independently), fix wave, ff-merged at `47087c6`.
- [x] Add the missing Phase 11b row to the master plan's phase table. `a99ab43`.

### Verification Plan
- On `main` after the merge: core `failed: 0` (total recorded in the ledger), Excel `93 / 0`, `dotnet csharpier check .` clean, Release build 0 warnings.
- `git log main -40` carries no `Co-Authored-By: Claude` / `Claude-Session` line.
- STOP HERE and tell the user: `main` is green, 3.20.0 can be cut (`release.yml`, `workflow_dispatch`). Phase B starts only after the user says so or cuts the release.

### Phase Summary

**Complete, merged at `47087c6`.** `main` is green for the first time since Phase 7 registered its producers: core **2008 / 0**, Excel **93 / 0**, csharpier clean, Release 0 warnings, no AI trailer in 45 commits. **3.20.0 can be cut** (`release.yml`, `workflow_dispatch`); `main` is 38 commits ahead of `origin/main` and needs a push first.

Four binding sites stopped collapsing a computed array to its top-left. The scope holds an `ArrayOperand` built once (`ArrayBindings.Capture`); `NameReference` answers that binding in each of its three roles; the one bare-reference predicate every gate shares became context-aware, and the context-free overload was removed for want of callers; `ResolveNameShape` gained an `Array` outcome; `Choose`, `UnaryOperation{Plus}` and `Let` gained probe/build arms. No new `ComputedValueKind`. Three Phase 8 pins and six Phase 11a pins flipped, each with both numbers in its commit body.

Five of the controller's own claims were measured false by the implementers and are worth carrying forward: a design item saying a gate "follows with no code of its own" was wrong twice (the criteria gate needed `ProbeLet`/`TryBuildLet`); item 11 as written would have created four new divergences, because making `+A1:A3` array-eligible at a top level makes the criteria gate refuse what the oracle reads (`COUNTIF(+A1:A3,">0")` is 2 there); item 10 needed CHOOSE's own capture rather than IF's; the scalar reading of a binding is the operand's top-left at every site, not `#VALUE!`; and `TryGetRegion` would have been a lie as a `Try*` returning three outcomes.

The review's most useful finding was not a defect but a coverage gap: a contributor-style revert of the predicate's defined-name clause left the suite fully green with four measured divergences silently restored. The unary-`+` half survived the same mutation through mechanism pins in two other files; the defined-name half had none and now does.

## Phase B: 4 — foundation task first, the rest in parallel

Status: Not started

Status corrected 2026-09-11: the worktree was created off `main` EARLY (tags 323-326 were already there), so Phase B ran concurrently with Phase A. It still merges AFTER Phase A.

Plan: `structured-table-references-and-aggregate/phase-4-lexer-parser.md` (re-verified 2026-09-10, three rulings). Briefs: `.superpowers/sdd/phase-4-lexer-parser/task-{1..7}-brief.md`. Worktree `/Volumes/Work/Develop/MySheet-p4`, branch `feat/lexer-parser`.

- [x] Task 1 — foundation: token, error kinds, `TableReference` with the six-member enum and tag 327, scanner. `d6e4062` + `3dbd6b2`, core 1927/1 → 1973/1, tag count 327 → 328.
- [x] Task 3 — lexer. `ed336c5`, `TokenizerTests` 9 → 17 cases, core → 1981/1.
- [x] Task 7 — name-validator repoint. `eaaa36d`, core → 1993/1.
- [x] Task 4 — anchored support. `e2761c2`, core → 2008/1.
- [ ] Task 2 — grammar + writer (running).
- [ ] Task 5 — parser arms (after 2 and 3).
- [ ] Task 6 — loader vehicle and every doc the phase falsifies (after 5).
- [ ] Two-part final review, fix wave, ff-merge.

### Verification Plan
- After Task 1's merge: core and Excel suites `failed: 0`; `grep -c '^\[MemoryPackUnion' Danfma.MySheet/Expressions/Expression.cs` = 328.
- After the phase: `FormulaWriterTests` 51 → 73, `TokenizerTests` 9 → 13, `ParseExceptionTests` 14 → 25, both suites 0 failures.

### Phase Summary
_(write when phase completes)_

## Phase C: 5 and 6 concurrently with Phase B's tail

Status: Not started

Plans: `phase-5-resolution-and-graph.md` (its items 1-4 are deleted by the 2026-09-10 ruling; it depends only on Phase B Task 1) and `phase-6-excel-loader.md`. Both need re-verification against `main` before briefs are generated — Phase 5's anchors are as old as Phase 4's were, and Phase 6 has never been audited.

- [x] Re-verify Phase 5 against `main`; rulings R1-R4 written into the phase file (`53010e4`); five briefs; worktree `/Volumes/Work/Develop/MySheet-p5`, branch `feat/resolution-and-graph` off `feat/lexer-parser`.
- [x] Re-verify Phase 6; nine rulings written into the phase file (`9bef4c7`); three briefs; worktree `/Volumes/Work/Develop/MySheet-p6`, branch `feat/excel-loader`.
- [x] Phase 6 Task 1 — the `<table>` reader, nine committed Aspose fixtures, export/merge pins. `ae6d82a`, Excel 93/0 → 125/0.
- [ ] Phase 5 Task 1 — six-area geometry plus the `Empty` outcome (running).
- [ ] Execute Phase 5 (its cadence), ff-merge.
- [ ] Execute Phase 6 (its cadence), ff-merge. Phase 6 is the release blocker: without it a real `.xlsx` cell holding `SUM(Tabela1[Valor])` answers `#NAME?`.

### Verification Plan
- Both suites 0 failures on `main` after each merge.
- An `.xlsx` produced by Aspose with a real table and `=SUM(Tabela1[Valor])` loads and answers the oracle's number (the end-to-end pin Phase 6 must add).

### Phase Summary
_(write when phase completes)_

## Phase D: release 3.21.0

Status: Not started

- [ ] Master plan rows for 4, 5, 6 marked Complete with counts.
- [ ] Tell the user `main` is green and 3.21.0 can be cut.

### Verification Plan
- Both suites 0 failures; no AI trailer in `git log`; `CHANGELOG.md` untouched (versionize owns it).

### Phase Summary
_(write when phase completes)_

## RESUME HERE (2026-09-11, end of session)

Four branches, none merged, every one clean at its last accepted commit. Merge order is fixed: **11c → main → cut 3.20.0 → p4 rebases onto main → p5 rebases onto p4 → p6 rebases last**.

| branch | worktree | head | suite there |
| --- | --- | --- | --- |
| `feat/array-bindings` | `MySheet-11c` | `1fd687d` | core 1988 / 15 — the 15 are Task 3's rows |
| `feat/lexer-parser` | `MySheet-p4` | `e2761c2` | core 2008 / 1 (11c's pin) |
| `feat/resolution-and-graph` | `MySheet-p5` | `eaaa36d` | core 1993 / 1 |
| `feat/excel-loader` | `MySheet-p6` | `ae6d82a` | Excel 125 / 0 |

Three tasks were in flight when the session ended — 11c T3, Phase 4 T2, Phase 5 T1. Check each worktree's `git status --short` first: if one is dirty, an agent died mid-edit and its brief plus the ledger say what it owed. Then dispatch what is missing, per-phase ledgers under `.superpowers/sdd/<phase>/progress.md`.

Next actions in order: finish 11c T3 (core to 0 failures), 11c T4 (docs, Sonnet), 11c's two-part review, ff-merge, **stop and tell the user 3.20.0 can be cut**. Then Phase 4 T2 → T5 → T6, Phase 5 T1 → T2/T3/T4 → T5, Phase 6 T2 → T3.

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete — releases are manual via `.github/workflows/release.yml` `workflow_dispatch`: versionize bumps and tags, packs both packages, publishes to NuGet.org via Trusted Publishing; the user runs it)_
