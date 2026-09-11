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

Status: In progress

Plan: `structured-table-references-and-aggregate/phase-11c-array-bindings.md`. Briefs: `.superpowers/sdd/phase-11c-array-bindings/task-{1..4}-brief.md`. Worktree `/Volumes/Work/Develop/MySheet-11c`, branch `feat/array-bindings`.

- [ ] Task 1 — the pins, red on arrival (items 1-3), plus the twelve unmeasured oracle rows.
- [ ] Task 2 — the scope carries an array; gates consult it (items 4-8).
- [ ] Task 3 — defined names, CHOOSE, unary `+` (items 9-12). Core suite reaches 0 failures here.
- [ ] Task 4 — docs, design of record, bookkeeping (items 13-15).
- [ ] Two-part final review, fix wave, ff-merge to `main`.
- [ ] Add the missing Phase 11b row to the master plan's phase table (found 2026-09-11).

### Verification Plan
- On `main` after the merge: core `failed: 0` (total recorded in the ledger), Excel `93 / 0`, `dotnet csharpier check .` clean, Release build 0 warnings.
- `git log main -40` carries no `Co-Authored-By: Claude` / `Claude-Session` line.
- STOP HERE and tell the user: `main` is green, 3.20.0 can be cut (`release.yml`, `workflow_dispatch`). Phase B starts only after the user says so or cuts the release.

### Phase Summary
_(write when phase completes)_

## Phase B: 4 — foundation task first, the rest in parallel

Status: Not started

Plan: `structured-table-references-and-aggregate/phase-4-lexer-parser.md` (re-verified 2026-09-10, three rulings). Briefs: `.superpowers/sdd/phase-4-lexer-parser/task-{1..7}-brief.md`. Worktree `/Volumes/Work/Develop/MySheet-p4`, branch `feat/lexer-parser`, created off `main` AFTER Phase A merges (the union tag is 327 only once Phase 7's 323-326 are on main — they are).

- [ ] Task 1 — foundation: token, error kinds, `TableReference` with the six-member enum and tag 327, scanner. Merge this task to `main` on its own as soon as it is green: it is all Phase 5 needs.
- [ ] Tasks 2, 3, 4, 7 concurrently (grammar+writer, lexer, anchored support, name-validator repoint).
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

- [ ] Re-verify Phase 5 against `main` (a read-only agent, same brief shape as Phase 4's re-verification); write its rulings; generate briefs; worktree `/Volumes/Work/Develop/MySheet-p5`.
- [ ] Re-verify Phase 6 the same way; briefs; worktree `/Volumes/Work/Develop/MySheet-p6`.
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

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete — releases are manual via `.github/workflows/release.yml` `workflow_dispatch`: versionize bumps and tags, packs both packages, publishes to NuGet.org via Trusted Publishing; the user runs it)_
