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
- [x] Task 2 — grammar + writer. `fd5fe81` + `40bcc93` + `f4217ed` + `f12e1c1` + `eefd87d` + `f53e367` (six commits, TDD red-first, both suites green at every boundary). core → 2369/0 on the final branch head.
- [x] Task 5 — parser arms (MERGED with T6, the reason is measured: the arms re-arm eleven `Load_*` tests, so both halves landed together). `621e11a` + `d224b80` (core 2369/0, Excel 93/11 → 93/0).
- [x] Task 6 — folded into T5 (loader vehicle + the eleven `Load_*` tests) and `a26cdbd` (the serialization.md tag sentences, fixed after the two reviews of the partial edit).
- [ ] Final review, fix wave, ff-merge. Part 1 (GLM-5.3 via z.ai, launched before the user's ruling) returned Yes-with-fixes, all registration-only; the controller verified all nine findings against the tree. Part 2 running on the controller's own subagent; part 3 running on Copilot. Kimi 403 (monthly). Reviewer lineup and the no-more-Claude rule recorded in RESUME HERE.

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
- [x] Phase 5 Task 1 — six-area geometry plus the `Empty` outcome. `5879a48`, core 1994/1 → **2092/1** (+98), Excel 93/0. `TryGetRegion` renamed to `GetRegion`, returns `TableRegionOutcome { Resolved, Empty, Absent }`.
- [ ] Execute Phase 5 (its cadence), ff-merge. T2 `0654194` and T3 `c43799c`+`9f681f7` accepted (core 2436/0); T4 accepted 2026-09-11 session 2 (four commits, core **2451/0**, Excel 93/0, ledger has the details); T5 dispatched after Phase 4 merges — Opus-class work running on GLM-5.3-Flash instead (see Phase B note), and the p5 rebase owes the two INDIRECT canary flips (p5 ledger).
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

## RESUME HERE (2026-09-11, session 3 — 3.20.0 shipped)

**3.20.0 shipped** (tag `v3.20.0` + `chore(release): 3.20.0` on `origin/main`). Remaining: Phases 4, 5 and 6 for 3.21.0. Merge order fixed: **p4 → main, p5 rebases onto main, p6 rebases last**. No release before Phase 6 — without it a real `.xlsx` cell holding `SUM(Tabela1[Valor])` answers `#NAME?`.

| branch | worktree | head | suite there | state |
| --- | --- | --- | --- | --- |
| `feat/lexer-parser` | `MySheet-p4` | `a26cdbd` | core 2369 / 0, Excel 93 / 0 | six tasks landed, clean. GLM-5.3 review complete (Yes with fixes, all registration; findings verified by the controller); second review running on a GLM-5.3-Flash subagent |
| `feat/resolution-and-graph` | `MySheet-p5` | `2b6536b` | core 2451 / 0, Excel 93 / 0 | T1-T4 accepted, clean, off p4's pre-parser-arms history. T5 waits on the p4 merge; the rebase owes two INDIRECT canary flips (p5 ledger) |
| `feat/excel-loader` | `MySheet-p6` | `ae6d82a` | Excel 125 / 0 | T1 landed, clean. T2 unblocked only after Phases 4+5 merge |

**Reviewer lineup (user, 2026-09-11 session 3): every final review = one instance of the controller's own subagent + Copilot. Kimi is out (403, monthly limit). Claude/z.ai dispatches are OFF until Monday Sep 14** (Opus+sonnet weekly limit; Fable credits; glm-5.3 was used for Phase 4's first part before the ruling and its findings stay as verified inputs, but no new z.ai/claude-CLI launches). Copilot is back by user decision after four phases of fabrication — every Copilot finding is verified against the tree before it enters a fix wave. Implementation tasks that were Opus-class run on the controller's own subagent instance.

Next actions in order: (1) de-dup the P4 reviews (GLM-5.3's completed part + the controller-subagent part + Copilot's part) into `phase-4-lexer-parser/fix-wave.md` (11c's file is the format), verify every new finding against the tree, execute the fixes as controller, `git rebase main` (conflicts expected: sweep file — main's `e7b50f0` item-34 edit vs the wave's new sweep item; master plan phase table — keep both sides' rows), `git merge --ff-only`; (2) rebase p5, flip the two INDIRECT canaries to 60.0 (or hand them to T5's brief — decided at the time), dispatch T5 on the controller's subagent instance (brief `task-5-brief.md`, "no numbered items by design" — scope is R3 + M4 pin + finding 8's docs); (3) Phase 5 review (controller subagent + Copilot), fix wave, ff-merge; (4) Phase 6 T2 then T3 (T3 re-measures every anchor — several are 3-8 lines off), review, fix wave, ff-merge; (5) master plan rows Complete, tell the user 3.21.0 can be cut.

## Final Recap
_(write when all phases complete)_

## Deployment Plan
_(write when all phases complete — releases are manual via `.github/workflows/release.yml` `workflow_dispatch`: versionize bumps and tags, packs both packages, publishes to NuGet.org via Trusted Publishing; the user runs it)_
